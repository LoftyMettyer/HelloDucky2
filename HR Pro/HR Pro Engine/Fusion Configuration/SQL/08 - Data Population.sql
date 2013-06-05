
DELETE FROM fusion.[MessageElements]
DELETE FROM fusion.[Message]

DELETE FROM [fusion].[Element]
DELETE FROM [fusion].[Category]

-- Categories
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (0, N'Employee', 1)
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (1, N'Salary', NULL)

-- Elements
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1, 0, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (2, 0, N'Staff Number', NULL, -1, NULL, NULL, 4, 0)
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (3, 0, N'Surname', NULL, -1, NULL, NULL, 2, 0)
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (4, 0, N'Forenames', NULL, -1, NULL, NULL, 3, 0)
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (5, 0, N'Title', NULL, -1, NULL, NULL, 13, 0)
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (6, 0, N'NI Number', NULL, -1, NULL, NULL, NULL, 0)
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (7, 0, N'Date of Birth', NULL, -1, NULL, NULL, NULL, 0)
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (8, 0, N'Gender', NULL, -1, NULL, NULL, 18, 0)
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (9, 0, N'Marital Status', NULL, -1, NULL, NULL, NULL, 0)
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (10, 0, N'Address Line 1', NULL, -1, NULL, NULL, 23, 0)
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (11, 0, N'Address Line 2', NULL, -1, NULL, NULL, 24, 0)
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (12, 0, N'Address Line 3', NULL, -1, NULL, NULL, 25, 0)
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (13, 0, N'Address Line 4', NULL, -1, NULL, NULL, NULL, 0)
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (14, 0, N'Address Line 5', NULL, -1, NULL, NULL, NULL, 0)
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (15, 0, N'Post Code', NULL, -1, NULL, NULL, NULL, 0)
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (16, 0, N'Telephone (Home)', NULL, -1, NULL, NULL, NULL, 0)
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (17, 0, N'Mobile (Personal)', NULL, -1, NULL, NULL, NULL, 0)
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (18, 0, N'Text Payment Advice', NULL, -1, NULL, NULL, NULL, 0)
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (19, 0, N'Email (Work)', NULL, -1, NULL, NULL, NULL, 0)
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (20, 0, N'Email Payslip', NULL, -1, NULL, NULL, NULL, 0)
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (21, 0, N'Start Date', NULL, -1, NULL, NULL, 14, 0)
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (22, 0, N'Leaving Date', NULL, -1, NULL, NULL, 15, 0)

INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (200, 1, N'Start Date', NULL, -1, NULL, NULL, NULL, 0)
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (201, 1, N'End Date', NULL, -1, NULL, NULL, NULL, 0)
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (202, 1, N'Amount1', NULL, -1, NULL, NULL, NULL, 0)
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (203, 1, N'Grade', NULL, -1, NULL, NULL, NULL, 0)

INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (23, 0, N'Leaving Reason', NULL, -1, NULL, NULL, NULL, 0)
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (24, 0, N'Email (Personal)', NULL, -1, NULL, NULL, NULL, 0)
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (25, 0, N'Telephone (Work)', NULL, -1, NULL, NULL, NULL, 0)
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (26, 0, N'Mobile (Work)', NULL, -1, NULL, NULL, NULL, 0)
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (27, 0, N'Job Title', NULL, -1, NULL, NULL, NULL, 0)




-- Messages (based on xsd definitions)
	INSERT fusion.[Message] (ID, Name, Description, [Schema], Skeleton, Version, AllowPublish, AllowSubscribe, Publish, Subscribe, StopDeletion, BypassValidation)
		VALUES (1, 'StaffChange', ' Change of staff details', 0x, 1, 1, 1, 1, 1, 1, 1, 1)

	INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
		VALUES (1, 4, 'forenames', 1, 12, 1, 1, 1, 20, 50, 0)

	INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
		VALUES (1, 3, 'surname', 2, 12, 1, 1, 1, 20, 50, 0)

	INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
		VALUES (1, 3, 'preferredName', 3, 12, 1, 1, 1, 20, 50, 0)

	INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
		VALUES (1, 2, 'payrollNumber', 4, 12, 1, 1, 1, 20, 50, 0)

	INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
		VALUES (1, 7, 'DOB', 5, 11, 1, 1, 1, 20, 50, 0)

	INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
		VALUES (1, 2, 'employeeType', 5, 12, 1, 1, 1, 20, 50, 0)

	--INSERT fusion.[MessageElements] (ID, MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
	--	VALUES (2, 1, 2, 'employmentStatus', 5, 12, 1, 1, 1, 20, 50, 0)

	INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
		VALUES (1, 16, 'homePhoneNumber', 5, 12, 1, 1, 1, 20, 50, 0)

	INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
		VALUES (1, 26, 'workMobile', 5, 12, 1, 1, 1, 20, 50, 0)

	INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
		VALUES (1, 17, 'personalMobile', 5, 12, 1, 1, 1, 20, 50, 0)

	INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
		VALUES (1, 10, 'addressLine1', 5, 12, 1, 1, 1, 20, 50, 0)

	INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
		VALUES (1, 11, 'addressLine2', 5, 12, 1, 1, 1, 20, 50, 0)

	INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
		VALUES (1, 12, 'addressLine3', 5, 12, 1, 1, 1, 20, 50, 0)

	INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
		VALUES (1, 13, 'addressLine4', 5, 12, 1, 1, 1, 20, 50, 0)

	INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
		VALUES (1, 14, 'addressLine5', 5, 12, 1, 1, 1, 20, 50, 0)

	INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
		VALUES (1, 15, 'postCode', 5, 12, 1, 1, 1, 20, 50, 0)

	INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
		VALUES (1, 8, 'gender', 5, 12, 1, 1, 1, 20, 50, 1)

	INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
		VALUES (1, 21, 'startDate', 5, 12, 1, 1, 1, 20, 50, 0)

	INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
		VALUES (1, 22, 'leavingDate', 5, 12, 1, 1, 1, 20, 50, 0)

	INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
		VALUES (1, 23, 'leavingReason', 5, 12, 1, 1, 1, 20, 50, 0)

	INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
		VALUES (1, 1, 'companyName', 5, 12, 1, 1, 1, 20, 50, 0)

	INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
		VALUES (1, 27, 'jobTitle', 5, 12, 1, 1, 1, 20, 50, 0)

/*

GO

INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (2, N'Allowances', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (3, N'Loans', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (4, N'Deductions', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (5, N'SSP', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (6, N'SMP', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (7, N'SPP', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (8, N'SAP', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (11, N'Extra Allowance - Accommodation', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (12, N'Extra Allowance - Benefits', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (13, N'Extra Allowance - Bonuses', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (14, N'Extra Allowance - Commissions', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (15, N'Extra Allowance - Holiday Sale', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (16, N'Extra Allowance - Insurance', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (17, N'Extra Allowance - Meals', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (18, N'Extra Allowance - Overtime', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (19, N'Extra Allowance - Pension', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (20, N'Extra Allowance - Travel', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (21, N'Extra Allowance - Vehicle', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (22, N'Extra Allowance - Weightings', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (23, N'Extra Allowance - User Defined 1', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (24, N'Extra Allowance - User Defined 2', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (25, N'Extra Allowance - User Defined 3', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (26, N'Extra Allowance - User Defined 4', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (27, N'Extra Allowance - User Defined 5', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (28, N'Extra Allowance - User Defined 6', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (29, N'Extra Allowance - User Defined 7', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (30, N'Extra Allowance - User Defined 8', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (31, N'Extra Allowance - User Defined 9', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (32, N'Extra Allowance - User Defined 10', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (41, N'Extra Deduction - Holiday Buy', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (42, N'Extra Deduction - User Defined 1', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (43, N'Extra Deduction - User Defined 2', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (44, N'Extra Deduction - User Defined 3', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (45, N'Extra Deduction - User Defined 4', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (46, N'Extra Deduction - User Defined 5', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (47, N'Extra Deduction - User Defined 6', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (48, N'Extra Deduction - User Defined 7', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (49, N'Extra Deduction - User Defined 8', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (50, N'Extra Deduction - User Defined 9', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (51, N'Extra Deduction - User Defined 10', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (61, N'Pay Scale Group', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (62, N'Pay Scale', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (63, N'Point', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (64, N'Post', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (65, N'Appointment', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (66, N'Negotiating Body', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (67, N'Post Status', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (68, N'Location', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (69, N'Duty Type', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (70, N'Appointment Information', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (71, N'Pension', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (72, N'Absence', 2)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (73, N'SPP Adoption', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (74, N'Working Pattern', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (75, N'Keeping in Touch Days (Maternity)', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (76, N'Keeping in Touch Days (Adoption)', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (77, N'ASPP Adoption', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (78, N'ASPP Birth', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (79, N'Keeping in Touch Days (ASPP Adoption)', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (80, N'Keeping in Touch Days (ASPP Birth)', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (101, N'Leave Reason', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (102, N'Marital Status', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (103, N'Cost Centre', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (104, N'Job Grade', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (105, N'Sort Code 1', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (106, N'Sort Code 2', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (107, N'Sort Code 3', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (108, N'Sort Code 4', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (109, N'Sort Code 5', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (110, N'Sort Code 6', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (111, N'Job Title', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (112, N'Ethnic Origin', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (113, N'Nationality', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (114, N'Reports To (1)', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (115, N'Reports To (2)', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (116, N'Absence Type', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (117, N'Absence Reason', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (118, N'Bank Details', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (131, N'Extra Code Table - User Defined 1', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (132, N'Extra Code Table - User Defined 2', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (133, N'Extra Code Table - User Defined 3', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (134, N'Extra Code Table - User Defined 4', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (135, N'Extra Code Table - User Defined 5', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (136, N'Extra Code Table - User Defined 6', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (137, N'Extra Code Table - User Defined 7', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (138, N'Extra Code Table - User Defined 8', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (139, N'Extra Code Table - User Defined 9', NULL)
GO
INSERT [fusion].[Category] ([ID], [Name], [TableID]) VALUES (140, N'Extra Code Table - User Defined 10', NULL)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1, 0, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (2, 0, N'Employee Code', NULL, -1, NULL, NULL, 4, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (3, 0, N'Employee Surname', NULL, -1, NULL, NULL, 2, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (4, 0, N'Employee Forenames', NULL, -1, NULL, NULL, 3, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (5, 0, N'Employee Title', NULL, -1, NULL, NULL, 13, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (6, 0, N'Employee NI Number', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (7, 0, N'Employee Date of Birth', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (8, 0, N'Employee Gender', NULL, -1, NULL, NULL, 18, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (9, 0, N'Employee Marital Status', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (10, 0, N'Employee Address Line 1', NULL, -1, NULL, NULL, 23, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (11, 0, N'Employee Address Line 2', NULL, -1, NULL, NULL, 24, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (12, 0, N'Employee Address Line 3', NULL, -1, NULL, NULL, 25, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (13, 0, N'Employee Address Line 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (14, 0, N'Employee Address Line 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (15, 0, N'Employee Address Post Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (16, 0, N'Telephone', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (17, 0, N'Mobile', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (18, 0, N'Text Payment Advice', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (19, 0, N'Email', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (20, 0, N'Email Payslip', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (21, 0, N'Employment Date', NULL, -1, NULL, NULL, 14, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (22, 0, N'Leaving Date', NULL, -1, NULL, NULL, 15, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (23, 0, N'Payment Frequency', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (24, 0, N'Payment Method', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (25, 0, N'Bank Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (26, 0, N'Branch Address Line 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (27, 0, N'Branch Address Line 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (28, 0, N'Branch Address Line 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (29, 0, N'Branch Address Line 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (30, 0, N'Branch Address Post Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (31, 0, N'Branch Sort Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (32, 0, N'Bank Account Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (33, 0, N'Bank Account Number', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (34, 0, N'BACS Reference Number', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (35, 0, N'Autopay Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (36, 0, N'Account Type', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (37, 0, N'Department Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (38, 0, N'Employee Category Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (39, 0, N'Nominal Costs Account', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (40, 0, N'Cost Code 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (41, 0, N'Cost Code 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (42, 0, N'Cost Code 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (43, 0, N'Cost Code 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (44, 0, N'Cost Code 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (45, 0, N'Cost Code 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (46, 0, N'Cost Code 7', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (47, 0, N'Cost Code 8', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (48, 0, N'Cost Code 9', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (49, 0, N'Tax Code + W1/M1 Basis', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (50, 0, N'P45 Previous Employment Taxable Pay', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (51, 0, N'P45 Previous Employment Tax Paid', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (52, 0, N'NI Letter', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (53, 0, N'Full Time Equivalent Hours', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (54, 0, N'Contracted Hours', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (55, 0, N'Part Timer Flag', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (56, 0, N'Director Flag', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (57, 0, N'Director Start Week', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (58, 0, N'Known As', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (59, 0, N'Additional Email', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (60, 0, N'Pension Scheme', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (61, 0, N'OMP Scheme', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (62, 0, N'P11d', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (63, 0, N'Personnel No', NULL, -1, NULL, NULL, 4, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (64, 0, N'Hours Per Day', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (65, 0, N'Hours Per Month', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (66, 0, N'Reports To (1)', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (67, 0, N'Reports To (2)', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (68, 0, N'Unused', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (69, 0, N'Unused', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (70, 0, N'Unused', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (71, 0, N'Unused', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (72, 0, N'Unused', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (73, 0, N'Unused', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (74, 0, N'Unused', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (75, 0, N'Unused', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (76, 0, N'Unused', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (77, 0, N'Unused', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (78, 0, N'Unused', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (79, 0, N'12 Months Rolling Sick Days', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (80, 0, N'Current Period Sick Days', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (81, 0, N'Car Engine Size', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (82, 0, N'Employment Type', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (83, 0, N'Car User Category', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (84, 0, N'OSP Contract Type', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (85, 0, N'Student Loan', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (86, 0, N'Salary Type', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (87, 0, N'Use Spinal Points Flag', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (88, 0, N'Payroll Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (89, 0, N'Starter Form Type', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (90, 0, N'Starter Form Status', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (91, 0, N'User Definable Amount 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (92, 0, N'User Definable Amount 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (93, 0, N'User Definable Amount 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (94, 0, N'User Definable Amount 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (95, 0, N'User Definable Amount 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (96, 0, N'User Definable Amount 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (97, 0, N'User Definable Amount 7', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (98, 0, N'User Definable Amount 8', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (99, 0, N'User Definable Amount 9', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (100, 0, N'User Definable Amount 10', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (101, 0, N'User Definable Amount 11', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (102, 0, N'User Definable Amount 12', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (103, 0, N'User Definable Amount 13', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (104, 0, N'User Definable Amount 14', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (105, 0, N'User Definable Amount 15', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (106, 0, N'User Definable Amount 16', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (107, 0, N'User Definable Amount 17', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (108, 0, N'User Definable Amount 18', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (109, 0, N'User Definable Amount 19', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (110, 0, N'User Definable Flag 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (111, 0, N'User Definable Flag 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (112, 0, N'User Definable Flag 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (113, 0, N'User Definable Flag 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (114, 0, N'User Definable Flag 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (115, 0, N'User Definable Flag 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (116, 0, N'User Definable Flag 7', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (117, 0, N'User Definable Flag 8', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (118, 0, N'User Definable Flag 9', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (119, 0, N'User Definable Flag 10', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (120, 0, N'User Definable Flag 11', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (121, 0, N'User Definable Flag 12', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (122, 0, N'User Definable Flag 13', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (123, 0, N'User Definable Flag 14', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (124, 0, N'User Definable Flag 15', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (125, 0, N'User Definable Flag 16', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (126, 0, N'User Definable Flag 17', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (127, 0, N'User Definable Flag 18', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (128, 0, N'User Definable Flag 19', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (129, 0, N'User Definable Date 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (130, 0, N'User Definable Date 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (131, 0, N'User Definable Date 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (132, 0, N'User Definable Date 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (133, 0, N'User Definable Date 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (134, 0, N'User Definable Date 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (135, 0, N'User Definable Date 7', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (136, 0, N'User Definable Date 8', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (137, 0, N'User Definable Date 9', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (138, 0, N'User Definable Date 10', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (139, 0, N'User Definable Date 11', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (140, 0, N'User Definable Date 12', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (141, 0, N'User Definable Date 13', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (142, 0, N'User Definable Date 14', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (143, 0, N'User Definable Date 15', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (144, 0, N'User Definable Date 16', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (145, 0, N'User Definable Date 17', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (146, 0, N'User Definable Date 19', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (147, 0, N'User Definable Date 19', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (148, 0, N'User Definable Text 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (149, 0, N'User Definable Text 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (150, 0, N'User Definable Text 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (151, 0, N'User Definable Text 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (152, 0, N'User Definable Text 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (153, 0, N'User Definable Text 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (154, 0, N'User Definable Text 7', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (155, 0, N'User Definable Text 8', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (156, 0, N'User Definable Text 9', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (157, 0, N'User Definable Text 10', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (158, 0, N'User Definable Text 11', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (159, 0, N'User Definable Text 12', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (160, 0, N'User Definable Text 13', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (161, 0, N'User Definable Text 14', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (162, 0, N'User Definable Text 15', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (163, 0, N'User Definable Text 16', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (164, 0, N'User Definable Text 17', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (165, 0, N'User Definable Text 18', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (166, 0, N'User Definable Text 19', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (167, 0, N'Department Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (168, 0, N'Employee Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (169, 0, N'Tax Code Effective From Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (170, 0, N'NI Letter Effective From Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (171, 0, N'Employee Category Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (172, 0, N'Nominal Costs Account Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (173, 0, N'Costs Code 1 Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (174, 0, N'Costs Code 2 Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (175, 0, N'Costs Code 3 Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (176, 0, N'Costs Code 4 Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (177, 0, N'Costs Code 5 Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (178, 0, N'Costs Code 6 Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (179, 0, N'Costs Code 7 Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (180, 0, N'Costs Code 8 Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (181, 0, N'Costs Code 9 Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (182, 0, N'Passport Number', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (183, 0, N'Seconded', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (184, 0, N'EPM6 (Modified) Scheme', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (185, 0, N'EEA/Commonwealth Citizen', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (186, 0, N'Starter Statement A', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (187, 0, N'Starter Statement B', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (188, 0, N'Starter Statement C', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (189, 0, N'Irregular Payment Indicator', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (190, 0, N'Student Loan Indicator', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (191, 0, N'Foreign Country', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (192, 0, N'Stay in UK for 6 months or more', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (193, 0, N'Stay in UK less than 6 Months', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (194, 0, N'Work both in/out UK but living abroad', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (195, 0, N'Pension paid because recently bereaved', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (196, 0, N'Annual Pension', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (197, 1, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (198, 1, N'Employee Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (199, 1, N'Contract No', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (200, 1, N'Start Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (201, 1, N'End Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (202, 1, N'Amount1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (203, 1, N'Grade', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (204, 1, N'Amount2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (205, 1, N'Nominal Cost Account', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (206, 1, N'Cost Code 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (207, 1, N'Cost Code 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (208, 1, N'Cost Code 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (209, 1, N'Cost Code 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (210, 1, N'Cost Code 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (211, 1, N'Cost Code 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (212, 1, N'Cost Code 7', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (213, 1, N'Cost Code 8', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (214, 1, N'Cost Code 9', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (215, 1, N'Post Id', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (216, 1, N'Employee Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (217, 1, N'Department Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (218, 1, N'Department Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (219, 1, N'Payroll Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (220, 2, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (221, 2, N'Employee Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (222, 2, N'Allowance Type', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (223, 2, N'Start Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (224, 2, N'End Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (225, 2, N'Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (226, 2, N'Nominal Account', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (227, 2, N'Cost Code 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (228, 2, N'Cost Code 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (229, 2, N'Cost Code 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (230, 2, N'Cost Code 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (231, 2, N'Cost Code 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (232, 2, N'Cost Code 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (233, 2, N'Cost Code 7', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (234, 2, N'Cost Code 8', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (235, 2, N'Cost Code 9', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (236, 2, N'Employee Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (237, 2, N'Department Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (238, 2, N'Department Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (239, 2, N'Payroll Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (240, 3, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (241, 3, N'Employee Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (242, 3, N'Loan Type', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (243, 3, N'Start Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (244, 3, N'End Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (245, 3, N'Period Repayment Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (246, 3, N'Outstanding Balance', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (247, 3, N'Repaid Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (248, 3, N'Reference', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (249, 3, N'Nominal Account', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (250, 3, N'Cost Code 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (251, 3, N'Cost Code 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (252, 3, N'Cost Code 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (253, 3, N'Cost Code 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (254, 3, N'Cost Code 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (255, 3, N'Cost Code 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (256, 3, N'Cost Code 7', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (257, 3, N'Cost Code 8', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (258, 3, N'Cost Code 9', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (259, 3, N'Employee Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (260, 3, N'Department Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (261, 3, N'Department Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (262, 3, N'Payroll Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (263, 4, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (264, 4, N'Employee Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (265, 4, N'Deduction Type', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (266, 4, N'Start Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (267, 4, N'End Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (268, 4, N'Deduction Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (269, 4, N'Reference', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (270, 4, N'Nominal Account', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (271, 4, N'Cost Code 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (272, 4, N'Cost Code 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (273, 4, N'Cost Code 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (274, 4, N'Cost Code 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (275, 4, N'Cost Code 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (276, 4, N'Cost Code 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (277, 4, N'Cost Code 7', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (278, 4, N'Cost Code 8', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (279, 4, N'Cost Code 9', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (280, 4, N'Employee Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (281, 4, N'Department Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (282, 4, N'Department Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (283, 4, N'Payroll Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (284, 5, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (285, 5, N'Employee Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (286, 5, N'Start Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (287, 5, N'Start Session', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (288, 5, N'End Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (289, 5, N'End Session', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (290, 5, N'Reference', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (291, 5, N'Nominal Account', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (292, 5, N'Cost Code 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (293, 5, N'Cost Code 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (294, 5, N'Cost Code 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (295, 5, N'Cost Code 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (296, 5, N'Cost Code 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (297, 5, N'Cost Code 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (298, 5, N'Cost Code 7', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (299, 5, N'Cost Code 8', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (300, 5, N'Cost Code 9', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (301, 5, N'Employee Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (302, 5, N'Department Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (303, 5, N'Department Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (304, 5, N'Payroll Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (305, 6, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (306, 6, N'Employee Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (307, 6, N'Expected Week of Confinement', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (308, 6, N'Start Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (309, 6, N'End Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (310, 6, N'Reference', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (311, 6, N'Nominal Account', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (312, 6, N'Cost Code 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (313, 6, N'Cost Code 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (314, 6, N'Cost Code 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (315, 6, N'Cost Code 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (316, 6, N'Cost Code 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (317, 6, N'Cost Code 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (318, 6, N'Cost Code 7', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (319, 6, N'Cost Code 8', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (320, 6, N'Cost Code 9', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (321, 6, N'Employee Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (322, 6, N'Department Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (323, 6, N'Department Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (324, 6, N'Payroll Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (325, 6, N'MATB1 Received Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (326, 6, N'Actual Date of Birth', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (327, 7, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (328, 7, N'Employee Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (329, 7, N'Expected Week of Confinement', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (330, 7, N'Baby Date of Birth', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (331, 7, N'SPP Start Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (332, 7, N'SPP Weeks Number', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (333, 7, N'End Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (334, 7, N'Reference', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (335, 7, N'Nominal Account', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (336, 7, N'Cost Code 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (337, 7, N'Cost Code 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (338, 7, N'Cost Code 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (339, 7, N'Cost Code 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (340, 7, N'Cost Code 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (341, 7, N'Cost Code 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (342, 7, N'Cost Code 7', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (343, 7, N'Cost Code 8', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (344, 7, N'Cost Code 9', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (345, 7, N'Employee Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (346, 7, N'Department Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (347, 7, N'Department Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (348, 7, N'Payroll Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (349, 7, N'SC3 Received Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (350, 7, N'Still Birth', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (351, 8, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (352, 8, N'Employee Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (353, 8, N'SAP Start Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (354, 8, N'SAP End Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (355, 8, N'Reference', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (356, 8, N'Nominal Account', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (357, 8, N'Cost Code 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (358, 8, N'Cost Code 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (359, 8, N'Cost Code 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (360, 8, N'Cost Code 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (361, 8, N'Cost Code 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (362, 8, N'Cost Code 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (363, 8, N'Cost Code 7', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (364, 8, N'Cost Code 8', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (365, 8, N'Cost Code 9', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (366, 8, N'Employee Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (367, 8, N'Department Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (368, 8, N'Department Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (369, 8, N'Payroll Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (370, 8, N'Matching Certificate Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (371, 8, N'Child Expected Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (372, 8, N'Actual Placed Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (373, 8, N'Work up to Placement', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (374, 11, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (375, 11, N'Employee Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (376, 11, N'Type', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (377, 11, N'Start Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (378, 11, N'End Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (379, 11, N'Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (380, 11, N'Nominal Cost Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (381, 11, N'Cost Code 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (382, 11, N'Cost Code 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (383, 11, N'Cost Code 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (384, 11, N'Cost Code 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (385, 11, N'Cost Code 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (386, 11, N'Cost Code 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (387, 11, N'Cost Code 7', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (388, 11, N'Cost Code 8', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (389, 11, N'Cost Code 9', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (390, 11, N'Employee Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (391, 11, N'Department Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (392, 11, N'Department Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (393, 11, N'Payroll Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (394, 12, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (395, 12, N'Employee Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (396, 12, N'Type', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (397, 12, N'Start Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (398, 12, N'End Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (399, 12, N'Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (400, 12, N'Nominal Cost Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (401, 12, N'Cost Code 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (402, 12, N'Cost Code 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (403, 12, N'Cost Code 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (404, 12, N'Cost Code 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (405, 12, N'Cost Code 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (406, 12, N'Cost Code 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (407, 12, N'Cost Code 7', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (408, 12, N'Cost Code 8', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (409, 12, N'Cost Code 9', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (410, 12, N'Employee Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (411, 12, N'Department Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (412, 12, N'Department Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (413, 12, N'Payroll Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (414, 13, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (415, 13, N'Employee Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (416, 13, N'Type', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (417, 13, N'Start Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (418, 13, N'End Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (419, 13, N'Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (420, 13, N'Nominal Cost Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (421, 13, N'Cost Code 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (422, 13, N'Cost Code 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (423, 13, N'Cost Code 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (424, 13, N'Cost Code 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (425, 13, N'Cost Code 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (426, 13, N'Cost Code 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (427, 13, N'Cost Code 7', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (428, 13, N'Cost Code 8', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (429, 13, N'Cost Code 9', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (430, 13, N'Employee Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (431, 13, N'Department Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (432, 13, N'Department Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (433, 13, N'Payroll Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (434, 14, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (435, 14, N'Employee Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (436, 14, N'Type', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (437, 14, N'Start Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (438, 14, N'End Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (439, 14, N'Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (440, 14, N'Nominal Cost Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (441, 14, N'Cost Code 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (442, 14, N'Cost Code 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (443, 14, N'Cost Code 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (444, 14, N'Cost Code 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (445, 14, N'Cost Code 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (446, 14, N'Cost Code 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (447, 14, N'Cost Code 7', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (448, 14, N'Cost Code 8', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (449, 14, N'Cost Code 9', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (450, 14, N'Employee Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (451, 14, N'Department Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (452, 14, N'Department Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (453, 14, N'Payroll Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (454, 15, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (455, 15, N'Employee Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (456, 15, N'Type', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (457, 15, N'Start Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (458, 15, N'End Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (459, 15, N'Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (460, 15, N'Nominal Cost Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (461, 15, N'Cost Code 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (462, 15, N'Cost Code 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (463, 15, N'Cost Code 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (464, 15, N'Cost Code 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (465, 15, N'Cost Code 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (466, 15, N'Cost Code 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (467, 15, N'Cost Code 7', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (468, 15, N'Cost Code 8', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (469, 15, N'Cost Code 9', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (470, 15, N'Employee Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (471, 15, N'Department Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (472, 15, N'Department Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (473, 15, N'Payroll Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (474, 16, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (475, 16, N'Employee Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (476, 16, N'Type', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (477, 16, N'Start Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (478, 16, N'End Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (479, 16, N'Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (480, 16, N'Nominal Cost Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (481, 16, N'Cost Code 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (482, 16, N'Cost Code 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (483, 16, N'Cost Code 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (484, 16, N'Cost Code 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (485, 16, N'Cost Code 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (486, 16, N'Cost Code 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (487, 16, N'Cost Code 7', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (488, 16, N'Cost Code 8', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (489, 16, N'Cost Code 9', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (490, 16, N'Employee Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (491, 16, N'Department Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (492, 16, N'Department Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (493, 16, N'Payroll Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (494, 17, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (495, 17, N'Employee Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (496, 17, N'Type', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (497, 17, N'Start Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (498, 17, N'End Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (499, 17, N'Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (500, 17, N'Nominal Cost Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (501, 17, N'Cost Code 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (502, 17, N'Cost Code 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (503, 17, N'Cost Code 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (504, 17, N'Cost Code 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (505, 17, N'Cost Code 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (506, 17, N'Cost Code 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (507, 17, N'Cost Code 7', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (508, 17, N'Cost Code 8', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (509, 17, N'Cost Code 9', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (510, 17, N'Employee Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (511, 17, N'Department Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (512, 17, N'Department Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (513, 17, N'Payroll Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (514, 18, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (515, 18, N'Employee Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (516, 18, N'Type', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (517, 18, N'Start Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (518, 18, N'End Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (519, 18, N'Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (520, 18, N'Nominal Cost Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (521, 18, N'Cost Code 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (522, 18, N'Cost Code 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (523, 18, N'Cost Code 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (524, 18, N'Cost Code 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (525, 18, N'Cost Code 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (526, 18, N'Cost Code 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (527, 18, N'Cost Code 7', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (528, 18, N'Cost Code 8', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (529, 18, N'Cost Code 9', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (530, 18, N'Employee Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (531, 18, N'Department Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (532, 18, N'Department Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (533, 18, N'Payroll Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (534, 19, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (535, 19, N'Employee Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (536, 19, N'Type', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (537, 19, N'Start Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (538, 19, N'End Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (539, 19, N'Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (540, 19, N'Nominal Cost Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (541, 19, N'Cost Code 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (542, 19, N'Cost Code 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (543, 19, N'Cost Code 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (544, 19, N'Cost Code 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (545, 19, N'Cost Code 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (546, 19, N'Cost Code 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (547, 19, N'Cost Code 7', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (548, 19, N'Cost Code 8', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (549, 19, N'Cost Code 9', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (550, 19, N'Employee Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (551, 19, N'Department Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (552, 19, N'Department Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (553, 19, N'Payroll Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (554, 20, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (555, 20, N'Employee Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (556, 20, N'Type', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (557, 20, N'Start Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (558, 20, N'End Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (559, 20, N'Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (560, 20, N'Nominal Cost Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (561, 20, N'Cost Code 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (562, 20, N'Cost Code 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (563, 20, N'Cost Code 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (564, 20, N'Cost Code 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (565, 20, N'Cost Code 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (566, 20, N'Cost Code 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (567, 20, N'Cost Code 7', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (568, 20, N'Cost Code 8', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (569, 20, N'Cost Code 9', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (570, 20, N'Employee Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (571, 20, N'Department Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (572, 20, N'Department Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (573, 20, N'Payroll Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (574, 21, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (575, 21, N'Employee Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (576, 21, N'Type', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (577, 21, N'Start Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (578, 21, N'End Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (579, 21, N'Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (580, 21, N'Nominal Cost Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (581, 21, N'Cost Code 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (582, 21, N'Cost Code 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (583, 21, N'Cost Code 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (584, 21, N'Cost Code 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (585, 21, N'Cost Code 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (586, 21, N'Cost Code 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (587, 21, N'Cost Code 7', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (588, 21, N'Cost Code 8', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (589, 21, N'Cost Code 9', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (590, 21, N'Employee Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (591, 21, N'Department Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (592, 21, N'Department Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (593, 21, N'Payroll Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (594, 22, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (595, 22, N'Employee Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (596, 22, N'Type', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (597, 22, N'Start Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (598, 22, N'End Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (599, 22, N'Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (600, 22, N'Nominal Cost Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (601, 22, N'Cost Code 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (602, 22, N'Cost Code 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (603, 22, N'Cost Code 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (604, 22, N'Cost Code 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (605, 22, N'Cost Code 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (606, 22, N'Cost Code 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (607, 22, N'Cost Code 7', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (608, 22, N'Cost Code 8', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (609, 22, N'Cost Code 9', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (610, 22, N'Employee Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (611, 22, N'Department Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (612, 22, N'Department Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (613, 22, N'Payroll Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (614, 23, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (615, 23, N'Employee Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (616, 23, N'Type', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (617, 23, N'Start Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (618, 23, N'End Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (619, 23, N'Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (620, 23, N'Nominal Cost Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (621, 23, N'Cost Code 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (622, 23, N'Cost Code 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (623, 23, N'Cost Code 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (624, 23, N'Cost Code 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (625, 23, N'Cost Code 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (626, 23, N'Cost Code 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (627, 23, N'Cost Code 7', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (628, 23, N'Cost Code 8', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (629, 23, N'Cost Code 9', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (630, 23, N'Employee Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (631, 23, N'Department Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (632, 23, N'Department Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (633, 23, N'Payroll Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (634, 24, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (635, 24, N'Employee Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (636, 24, N'Type', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (637, 24, N'Start Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (638, 24, N'End Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (639, 24, N'Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (640, 24, N'Nominal Cost Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (641, 24, N'Cost Code 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (642, 24, N'Cost Code 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (643, 24, N'Cost Code 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (644, 24, N'Cost Code 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (645, 24, N'Cost Code 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (646, 24, N'Cost Code 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (647, 24, N'Cost Code 7', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (648, 24, N'Cost Code 8', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (649, 24, N'Cost Code 9', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (650, 24, N'Employee Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (651, 24, N'Department Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (652, 24, N'Department Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (653, 24, N'Payroll Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (654, 25, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (655, 25, N'Employee Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (656, 25, N'Type', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (657, 25, N'Start Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (658, 25, N'End Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (659, 25, N'Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (660, 25, N'Nominal Cost Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (661, 25, N'Cost Code 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (662, 25, N'Cost Code 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (663, 25, N'Cost Code 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (664, 25, N'Cost Code 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (665, 25, N'Cost Code 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (666, 25, N'Cost Code 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (667, 25, N'Cost Code 7', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (668, 25, N'Cost Code 8', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (669, 25, N'Cost Code 9', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (670, 25, N'Employee Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (671, 25, N'Department Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (672, 25, N'Department Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (673, 25, N'Payroll Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (674, 26, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (675, 26, N'Employee Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (676, 26, N'Type', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (677, 26, N'Start Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (678, 26, N'End Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (679, 26, N'Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (680, 26, N'Nominal Cost Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (681, 26, N'Cost Code 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (682, 26, N'Cost Code 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (683, 26, N'Cost Code 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (684, 26, N'Cost Code 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (685, 26, N'Cost Code 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (686, 26, N'Cost Code 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (687, 26, N'Cost Code 7', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (688, 26, N'Cost Code 8', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (689, 26, N'Cost Code 9', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (690, 26, N'Employee Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (691, 26, N'Department Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (692, 26, N'Department Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (693, 26, N'Payroll Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (694, 27, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (695, 27, N'Employee Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (696, 27, N'Type', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (697, 27, N'Start Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (698, 27, N'End Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (699, 27, N'Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (700, 27, N'Nominal Cost Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (701, 27, N'Cost Code 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (702, 27, N'Cost Code 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (703, 27, N'Cost Code 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (704, 27, N'Cost Code 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (705, 27, N'Cost Code 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (706, 27, N'Cost Code 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (707, 27, N'Cost Code 7', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (708, 27, N'Cost Code 8', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (709, 27, N'Cost Code 9', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (710, 27, N'Employee Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (711, 27, N'Department Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (712, 27, N'Department Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (713, 27, N'Payroll Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (714, 28, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (715, 28, N'Employee Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (716, 28, N'Type', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (717, 28, N'Start Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (718, 28, N'End Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (719, 28, N'Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (720, 28, N'Nominal Cost Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (721, 28, N'Cost Code 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (722, 28, N'Cost Code 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (723, 28, N'Cost Code 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (724, 28, N'Cost Code 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (725, 28, N'Cost Code 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (726, 28, N'Cost Code 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (727, 28, N'Cost Code 7', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (728, 28, N'Cost Code 8', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (729, 28, N'Cost Code 9', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (730, 28, N'Employee Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (731, 28, N'Department Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (732, 28, N'Department Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (733, 28, N'Payroll Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (734, 29, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (735, 29, N'Employee Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (736, 29, N'Type', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (737, 29, N'Start Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (738, 29, N'End Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (739, 29, N'Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (740, 29, N'Nominal Cost Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (741, 29, N'Cost Code 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (742, 29, N'Cost Code 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (743, 29, N'Cost Code 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (744, 29, N'Cost Code 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (745, 29, N'Cost Code 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (746, 29, N'Cost Code 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (747, 29, N'Cost Code 7', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (748, 29, N'Cost Code 8', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (749, 29, N'Cost Code 9', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (750, 29, N'Employee Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (751, 29, N'Department Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (752, 29, N'Department Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (753, 29, N'Payroll Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (754, 30, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (755, 30, N'Employee Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (756, 30, N'Type', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (757, 30, N'Start Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (758, 30, N'End Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (759, 30, N'Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (760, 30, N'Nominal Cost Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (761, 30, N'Cost Code 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (762, 30, N'Cost Code 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (763, 30, N'Cost Code 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (764, 30, N'Cost Code 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (765, 30, N'Cost Code 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (766, 30, N'Cost Code 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (767, 30, N'Cost Code 7', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (768, 30, N'Cost Code 8', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (769, 30, N'Cost Code 9', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (770, 30, N'Employee Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (771, 30, N'Department Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (772, 30, N'Department Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (773, 30, N'Payroll Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (774, 31, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (775, 31, N'Employee Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (776, 31, N'Type', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (777, 31, N'Start Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (778, 31, N'End Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (779, 31, N'Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (780, 31, N'Nominal Cost Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (781, 31, N'Cost Code 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (782, 31, N'Cost Code 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (783, 31, N'Cost Code 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (784, 31, N'Cost Code 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (785, 31, N'Cost Code 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (786, 31, N'Cost Code 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (787, 31, N'Cost Code 7', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (788, 31, N'Cost Code 8', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (789, 31, N'Cost Code 9', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (790, 31, N'Employee Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (791, 31, N'Department Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (792, 31, N'Department Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (793, 31, N'Payroll Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (794, 32, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (795, 32, N'Employee Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (796, 32, N'Type', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (797, 32, N'Start Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (798, 32, N'End Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (799, 32, N'Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (800, 32, N'Nominal Cost Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (801, 32, N'Cost Code 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (802, 32, N'Cost Code 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (803, 32, N'Cost Code 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (804, 32, N'Cost Code 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (805, 32, N'Cost Code 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (806, 32, N'Cost Code 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (807, 32, N'Cost Code 7', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (808, 32, N'Cost Code 8', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (809, 32, N'Cost Code 9', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (810, 32, N'Employee Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (811, 32, N'Department Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (812, 32, N'Department Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (813, 32, N'Payroll Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (814, 41, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (815, 41, N'Employee Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (816, 41, N'Type', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (817, 41, N'Start Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (818, 41, N'End Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (819, 41, N'Deduction Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (820, 41, N'Reference', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (821, 41, N'Nominal Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (822, 41, N'Cost Code 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (823, 41, N'Cost Code 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (824, 41, N'Cost Code 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (825, 41, N'Cost Code 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (826, 41, N'Cost Code 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (827, 41, N'Cost Code 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (828, 41, N'Cost Code 7', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (829, 41, N'Cost Code 8', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (830, 41, N'Cost Code 9', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (831, 41, N'Employee Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (832, 41, N'Department Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (833, 41, N'Department Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (834, 41, N'Payroll Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (835, 42, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (836, 42, N'Employee Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (837, 42, N'Type', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (838, 42, N'Start Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (839, 42, N'End Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (840, 42, N'Deduction Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (841, 42, N'Reference', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (842, 42, N'Nominal Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (843, 42, N'Cost Code 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (844, 42, N'Cost Code 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (845, 42, N'Cost Code 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (846, 42, N'Cost Code 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (847, 42, N'Cost Code 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (848, 42, N'Cost Code 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (849, 42, N'Cost Code 7', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (850, 42, N'Cost Code 8', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (851, 42, N'Cost Code 9', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (852, 42, N'Employee Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (853, 42, N'Department Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (854, 42, N'Department Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (855, 42, N'Payroll Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (856, 43, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (857, 43, N'Employee Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (858, 43, N'Type', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (859, 43, N'Start Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (860, 43, N'End Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (861, 43, N'Deduction Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (862, 43, N'Reference', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (863, 43, N'Nominal Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (864, 43, N'Cost Code 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (865, 43, N'Cost Code 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (866, 43, N'Cost Code 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (867, 43, N'Cost Code 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (868, 43, N'Cost Code 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (869, 43, N'Cost Code 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (870, 43, N'Cost Code 7', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (871, 43, N'Cost Code 8', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (872, 43, N'Cost Code 9', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (873, 43, N'Employee Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (874, 43, N'Department Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (875, 43, N'Department Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (876, 43, N'Payroll Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (877, 44, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (878, 44, N'Employee Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (879, 44, N'Type', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (880, 44, N'Start Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (881, 44, N'End Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (882, 44, N'Deduction Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (883, 44, N'Reference', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (884, 44, N'Nominal Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (885, 44, N'Cost Code 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (886, 44, N'Cost Code 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (887, 44, N'Cost Code 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (888, 44, N'Cost Code 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (889, 44, N'Cost Code 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (890, 44, N'Cost Code 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (891, 44, N'Cost Code 7', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (892, 44, N'Cost Code 8', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (893, 44, N'Cost Code 9', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (894, 44, N'Employee Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (895, 44, N'Department Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (896, 44, N'Department Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (897, 44, N'Payroll Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (898, 45, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (899, 45, N'Employee Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (900, 45, N'Type', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (901, 45, N'Start Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (902, 45, N'End Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (903, 45, N'Deduction Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (904, 45, N'Reference', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (905, 45, N'Nominal Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (906, 45, N'Cost Code 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (907, 45, N'Cost Code 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (908, 45, N'Cost Code 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (909, 45, N'Cost Code 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (910, 45, N'Cost Code 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (911, 45, N'Cost Code 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (912, 45, N'Cost Code 7', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (913, 45, N'Cost Code 8', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (914, 45, N'Cost Code 9', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (915, 45, N'Employee Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (916, 45, N'Department Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (917, 45, N'Department Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (918, 45, N'Payroll Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (919, 46, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (920, 46, N'Employee Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (921, 46, N'Type', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (922, 46, N'Start Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (923, 46, N'End Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (924, 46, N'Deduction Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (925, 46, N'Reference', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (926, 46, N'Nominal Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (927, 46, N'Cost Code 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (928, 46, N'Cost Code 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (929, 46, N'Cost Code 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (930, 46, N'Cost Code 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (931, 46, N'Cost Code 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (932, 46, N'Cost Code 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (933, 46, N'Cost Code 7', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (934, 46, N'Cost Code 8', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (935, 46, N'Cost Code 9', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (936, 46, N'Employee Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (937, 46, N'Department Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (938, 46, N'Department Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (939, 46, N'Payroll Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (940, 47, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (941, 47, N'Employee Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (942, 47, N'Type', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (943, 47, N'Start Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (944, 47, N'End Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (945, 47, N'Deduction Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (946, 47, N'Reference', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (947, 47, N'Nominal Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (948, 47, N'Cost Code 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (949, 47, N'Cost Code 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (950, 47, N'Cost Code 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (951, 47, N'Cost Code 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (952, 47, N'Cost Code 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (953, 47, N'Cost Code 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (954, 47, N'Cost Code 7', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (955, 47, N'Cost Code 8', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (956, 47, N'Cost Code 9', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (957, 47, N'Employee Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (958, 47, N'Department Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (959, 47, N'Department Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (960, 47, N'Payroll Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (961, 48, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (962, 48, N'Employee Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (963, 48, N'Type', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (964, 48, N'Start Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (965, 48, N'End Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (966, 48, N'Deduction Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (967, 48, N'Reference', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (968, 48, N'Nominal Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (969, 48, N'Cost Code 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (970, 48, N'Cost Code 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (971, 48, N'Cost Code 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (972, 48, N'Cost Code 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (973, 48, N'Cost Code 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (974, 48, N'Cost Code 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (975, 48, N'Cost Code 7', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (976, 48, N'Cost Code 8', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (977, 48, N'Cost Code 9', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (978, 48, N'Employee Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (979, 48, N'Department Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (980, 48, N'Department Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (981, 48, N'Payroll Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (982, 49, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (983, 49, N'Employee Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (984, 49, N'Type', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (985, 49, N'Start Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (986, 49, N'End Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (987, 49, N'Deduction Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (988, 49, N'Reference', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (989, 49, N'Nominal Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (990, 49, N'Cost Code 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (991, 49, N'Cost Code 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (992, 49, N'Cost Code 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (993, 49, N'Cost Code 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (994, 49, N'Cost Code 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (995, 49, N'Cost Code 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (996, 49, N'Cost Code 7', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (997, 49, N'Cost Code 8', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (998, 49, N'Cost Code 9', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (999, 49, N'Employee Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1000, 49, N'Department Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1001, 49, N'Department Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1002, 49, N'Payroll Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1003, 50, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1004, 50, N'Employee Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1005, 50, N'Type', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1006, 50, N'Start Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1007, 50, N'End Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1008, 50, N'Deduction Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1009, 50, N'Reference', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1010, 50, N'Nominal Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1011, 50, N'Cost Code 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1012, 50, N'Cost Code 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1013, 50, N'Cost Code 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1014, 50, N'Cost Code 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1015, 50, N'Cost Code 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1016, 50, N'Cost Code 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1017, 50, N'Cost Code 7', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1018, 50, N'Cost Code 8', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1019, 50, N'Cost Code 9', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1020, 50, N'Employee Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1021, 50, N'Department Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1022, 50, N'Department Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1023, 50, N'Payroll Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1024, 51, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1025, 51, N'Employee Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1026, 51, N'Type', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1027, 51, N'Start Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1028, 51, N'End Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1029, 51, N'Deduction Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1030, 51, N'Reference', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1031, 51, N'Nominal Amount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1032, 51, N'Cost Code 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1033, 51, N'Cost Code 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1034, 51, N'Cost Code 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1035, 51, N'Cost Code 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1036, 51, N'Cost Code 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1037, 51, N'Cost Code 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1038, 51, N'Cost Code 7', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1039, 51, N'Cost Code 8', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1040, 51, N'Cost Code 9', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1041, 51, N'Employee Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1042, 51, N'Department Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1043, 51, N'Department Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1044, 51, N'Payroll Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1045, 61, N'Pay Scale Group', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1046, 61, N'Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1047, 61, N'Effective Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1048, 61, N'Increment Type', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1049, 61, N'Increment Cut Off Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1050, 61, N'Increment Due Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1051, 61, N'Increment Period', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1052, 61, N'Auto Step New Start', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1053, 61, N'Auto Step', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1054, 61, N'Payment Level', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1055, 61, N'Weekly Payslip Display', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1056, 61, N'Negotiating Body', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1057, 61, N'Hours per Week', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1058, 62, N'Pay Scale Group', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1059, 62, N'Pay Scale', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1060, 62, N'Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1061, 62, N'Effective Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1062, 62, N'Minimum Point', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1063, 62, N'Maximum Point', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1064, 62, N'Bar Point', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1065, 63, N'Pay Scale Group', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1066, 63, N'Pay Scale', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1067, 63, N'Point', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1068, 63, N'Effective Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1069, 63, N'Annual', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1070, 63, N'Monthly', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1071, 63, N'Weekly', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1072, 63, N'Hourly', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1073, 64, N'Post ID', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1074, 64, N'Post Title', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1075, 64, N'Effective Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1076, 64, N'Pay Scale Group', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1077, 64, N'Pay Scale', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1078, 64, N'Minimum Point', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1079, 64, N'Maximum Point', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1080, 64, N'Bar Point', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1081, 64, N'Contract Type', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1082, 64, N'Full or Part Time', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1083, 64, N'Post End Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1084, 64, N'In Use', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1085, 64, N'Cost Centre', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1086, 64, N'Reports To', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1087, 64, N'Post Status', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1088, 64, N'Location', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1089, 64, N'Duty Type', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1090, 64, N'Budget FTE', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1091, 64, N'Budget Headcount', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1092, 64, N'Budget Cost', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1093, 65, N'Post ID', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1094, 65, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1095, 65, N'Staff Number', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1096, 65, N'Effective Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1097, 65, N'Point', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1098, 65, N'Primary Job', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1099, 65, N'Protected Group', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1100, 65, N'Protected Scale', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1101, 65, N'Protected Point', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1102, 65, N'Appointment Reason', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1103, 65, N'Appointment Information', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1104, 65, N'Auto Increment', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1105, 65, N'Hours per Week', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1106, 65, N'Contract Type', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1107, 65, N'Full or Part Time', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1108, 65, N'Appointment End Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1109, 65, N'Next Review Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1110, 66, N'Code Table ID', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1111, 66, N'Negotiating Body', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1112, 66, N'Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1113, 66, N'Short Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1114, 66, N'Email Address', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1115, 66, N'Supplementary Field 1a', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1116, 66, N'Supplementary Field 1b', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1117, 66, N'Supplementary Field 1c', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1118, 66, N'Supplementary Field 1d', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1119, 66, N'Supplementary Field 1e', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1120, 66, N'Supplementary Field 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1121, 66, N'Supplementary Field 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1122, 66, N'Supplementary Field 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1123, 66, N'Supplementary Field 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1124, 66, N'Supplementary Field 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1125, 67, N'Code Table ID', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1126, 67, N'Post Status', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1127, 67, N'Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1128, 67, N'Short Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1129, 68, N'Code Table ID', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1130, 68, N'Location', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1131, 68, N'Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1132, 68, N'Short Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1133, 68, N'Email Address', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1134, 68, N'Supplementary Field 1a', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1135, 68, N'Supplementary Field 1b', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1136, 68, N'Supplementary Field 1c', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1137, 68, N'Supplementary Field 1d', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1138, 68, N'Supplementary Field 1e', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1139, 68, N'Supplementary Field 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1140, 68, N'Supplementary Field 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1141, 68, N'Supplementary Field 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1142, 68, N'Supplementary Field 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1143, 68, N'Supplementary Field 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1144, 69, N'Code Table ID', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1145, 69, N'Duty Type', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1146, 69, N'Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1147, 69, N'Short Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1148, 70, N'Code Table ID', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1149, 70, N'Appointment Information', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1150, 70, N'Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1151, 70, N'Short Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1152, 71, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1153, 71, N'Employee Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1154, 71, N'Pension Scheme No', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1155, 71, N'Pension Employee', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1156, 71, N'Pension Employer', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1157, 71, N'Pension AVC', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1158, 71, N'Pension Joining Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1159, 71, N'Pension Leaving Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1160, 71, N'Pension Policy No', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1161, 71, N'Employee Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1162, 71, N'Department Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1163, 71, N'Department Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1164, 71, N'Payroll Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1165, 72, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1166, 72, N'Employee Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1167, 72, N'Absence Type', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1168, 72, N'Absence Reason', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1169, 72, N'Start Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1170, 72, N'End Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1171, 72, N'Start Session', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1172, 72, N'End Session', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1173, 72, N'Hours Duration', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1174, 72, N'Start Time', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1175, 72, N'End Time', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1176, 72, N'Absence ID', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1177, 72, N'Memo', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1178, 73, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1179, 73, N'Employee Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1180, 73, N'SC4 Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1181, 73, N'Child Expected Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1182, 73, N'Actual Placed Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1183, 73, N'Start SSP Leave', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1184, 73, N'Work up to Placement', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1185, 74, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1186, 74, N'Employee Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1187, 74, N'Effective Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1188, 74, N'Working Pattern AM', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1189, 74, N'Working Pattern PM', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1190, 75, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1191, 75, N'Employee Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1192, 75, N'MATB1 Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1193, 75, N'Start Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1194, 75, N'End Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1195, 75, N'Reason', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1196, 76, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1197, 76, N'Employee Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1198, 76, N'Child Expected Date Adoption', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1199, 76, N'Start Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1200, 76, N'End Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1201, 76, N'Reason', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1202, 77, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1203, 77, N'Employee Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1204, 77, N'SC8 Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1205, 77, N'Notified Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1206, 77, N'Intended Start Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1207, 77, N'ASPP Start Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1208, 77, N'ASPP End Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1209, 77, N'Adopter SAP Start Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1210, 77, N'Actual Placed Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1211, 77, N'Return to Work Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1212, 77, N'Adopter Surname', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1213, 77, N'Adopter Address 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1214, 77, N'Adopter Address 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1215, 77, N'Adopter Address 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1216, 77, N'Adopter Address 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1217, 77, N'Adopter Address 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1218, 77, N'Adopter NI Number', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1219, 77, N'Adopter Employer Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1220, 77, N'Adopter Employer Address 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1221, 77, N'Adopter Employer Address 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1222, 77, N'Adopter Employer Address 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1223, 77, N'Adopter Employer Address 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1224, 77, N'Adopter Employer Address 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1225, 77, N'Date of Death of Adopter', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1226, 77, N'Adoption Form Confirmed', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1227, 77, N'Adopter Forename 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1228, 77, N'Adopter Forename 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1229, 78, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1230, 78, N'Employee Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1231, 78, N'SC7 Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1232, 78, N'Notified Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1233, 78, N'Intended Start Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1234, 78, N'Actual Birth Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1235, 78, N'ASPP Start Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1236, 78, N'ASPP End Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1237, 78, N'Mother MPP STart Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1238, 78, N'Return to Work Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1239, 78, N'Mother Surname', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1240, 78, N'Mother Address 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1241, 78, N'Mother Address 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1242, 78, N'Mother Address 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1243, 78, N'Mother Address 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1244, 78, N'Mother Address 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1245, 78, N'Mother NI Number', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1246, 78, N'Mother Employer Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1247, 78, N'Mother Employer Address 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1248, 78, N'Mother Employer Address 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1249, 78, N'Mother Employer Address 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1250, 78, N'Mother Employer Address 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1251, 78, N'Mother Employer Address 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1252, 78, N'Date of Death of Mother', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1253, 78, N'Birth Certificate Confirmed', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1254, 78, N'Mother Forename 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1255, 78, N'Mother Forename 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1256, 79, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1257, 79, N'Employee Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1258, 79, N'SC8 Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1259, 79, N'Start Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1260, 79, N'End Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1261, 79, N'Reason', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1262, 80, N'Company Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1263, 80, N'Employee Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1264, 80, N'SC7 Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1265, 80, N'Start Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1266, 80, N'End Date', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1267, 80, N'Reason', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1268, 101, N'Code Table ID', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1269, 101, N'Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1270, 101, N'Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1271, 101, N'Short Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1272, 101, N'Email Address', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1273, 101, N'Supplementary Field 1a', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1274, 101, N'Supplementary Field 1b', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1275, 101, N'Supplementary Field 1c', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1276, 101, N'Supplementary Field 1d', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1277, 101, N'Supplementary Field 1e', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1278, 101, N'Supplementary Field 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1279, 101, N'Supplementary Field 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1280, 101, N'Supplementary Field 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1281, 101, N'Supplementary Field 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1282, 101, N'Supplementary Field 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1283, 102, N'Code Table ID', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1284, 102, N'Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1285, 102, N'Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1286, 102, N'Short Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1287, 102, N'Email Address', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1288, 102, N'Supplementary Field 1a', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1289, 102, N'Supplementary Field 1b', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1290, 102, N'Supplementary Field 1c', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1291, 102, N'Supplementary Field 1d', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1292, 102, N'Supplementary Field 1e', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1293, 102, N'Supplementary Field 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1294, 102, N'Supplementary Field 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1295, 102, N'Supplementary Field 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1296, 102, N'Supplementary Field 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1297, 102, N'Supplementary Field 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1298, 103, N'Code Table ID', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1299, 103, N'Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1300, 103, N'Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1301, 103, N'Short Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1302, 103, N'Email Address', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1303, 103, N'Supplementary Field 1a', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1304, 103, N'Supplementary Field 1b', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1305, 103, N'Supplementary Field 1c', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1306, 103, N'Supplementary Field 1d', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1307, 103, N'Supplementary Field 1e', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1308, 103, N'Supplementary Field 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1309, 103, N'Supplementary Field 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1310, 103, N'Supplementary Field 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1311, 103, N'Supplementary Field 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1312, 103, N'Supplementary Field 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1313, 103, N'Project', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1314, 104, N'Code Table ID', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1315, 104, N'Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1316, 104, N'Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1317, 104, N'Short Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1318, 104, N'Email Address', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1319, 104, N'Supplementary Field 1a', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1320, 104, N'Supplementary Field 1b', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1321, 104, N'Supplementary Field 1c', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1322, 104, N'Supplementary Field 1d', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1323, 104, N'Supplementary Field 1e', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1324, 104, N'Supplementary Field 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1325, 104, N'Supplementary Field 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1326, 104, N'Supplementary Field 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1327, 104, N'Supplementary Field 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1328, 104, N'Supplementary Field 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1329, 105, N'Code Table ID', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1330, 105, N'Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1331, 105, N'Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1332, 105, N'Short Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1333, 105, N'Email Address', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1334, 105, N'Supplementary Field 1a', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1335, 105, N'Supplementary Field 1b', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1336, 105, N'Supplementary Field 1c', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1337, 105, N'Supplementary Field 1d', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1338, 105, N'Supplementary Field 1e', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1339, 105, N'Supplementary Field 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1340, 105, N'Supplementary Field 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1341, 105, N'Supplementary Field 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1342, 105, N'Supplementary Field 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1343, 105, N'Supplementary Field 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1344, 106, N'Code Table ID', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1345, 106, N'Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1346, 106, N'Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1347, 106, N'Short Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1348, 106, N'Email Address', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1349, 106, N'Supplementary Field 1a', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1350, 106, N'Supplementary Field 1b', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1351, 106, N'Supplementary Field 1c', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1352, 106, N'Supplementary Field 1d', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1353, 106, N'Supplementary Field 1e', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1354, 106, N'Supplementary Field 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1355, 106, N'Supplementary Field 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1356, 106, N'Supplementary Field 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1357, 106, N'Supplementary Field 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1358, 106, N'Supplementary Field 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1359, 107, N'Code Table ID', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1360, 107, N'Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1361, 107, N'Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1362, 107, N'Short Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1363, 107, N'Email Address', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1364, 107, N'Supplementary Field 1a', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1365, 107, N'Supplementary Field 1b', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1366, 107, N'Supplementary Field 1c', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1367, 107, N'Supplementary Field 1d', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1368, 107, N'Supplementary Field 1e', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1369, 107, N'Supplementary Field 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1370, 107, N'Supplementary Field 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1371, 107, N'Supplementary Field 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1372, 107, N'Supplementary Field 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1373, 107, N'Supplementary Field 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1374, 108, N'Code Table ID', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1375, 108, N'Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1376, 108, N'Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1377, 108, N'Short Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1378, 108, N'Email Address', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1379, 108, N'Supplementary Field 1a', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1380, 108, N'Supplementary Field 1b', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1381, 108, N'Supplementary Field 1c', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1382, 108, N'Supplementary Field 1d', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1383, 108, N'Supplementary Field 1e', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1384, 108, N'Supplementary Field 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1385, 108, N'Supplementary Field 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1386, 108, N'Supplementary Field 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1387, 108, N'Supplementary Field 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1388, 108, N'Supplementary Field 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1389, 109, N'Code Table ID', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1390, 109, N'Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1391, 109, N'Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1392, 109, N'Short Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1393, 109, N'Email Address', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1394, 109, N'Supplementary Field 1a', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1395, 109, N'Supplementary Field 1b', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1396, 109, N'Supplementary Field 1c', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1397, 109, N'Supplementary Field 1d', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1398, 109, N'Supplementary Field 1e', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1399, 109, N'Supplementary Field 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1400, 109, N'Supplementary Field 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1401, 109, N'Supplementary Field 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1402, 109, N'Supplementary Field 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1403, 109, N'Supplementary Field 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1404, 110, N'Code Table ID', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1405, 110, N'Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1406, 110, N'Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1407, 110, N'Short Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1408, 110, N'Email Address', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1409, 110, N'Supplementary Field 1a', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1410, 110, N'Supplementary Field 1b', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1411, 110, N'Supplementary Field 1c', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1412, 110, N'Supplementary Field 1d', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1413, 110, N'Supplementary Field 1e', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1414, 110, N'Supplementary Field 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1415, 110, N'Supplementary Field 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1416, 110, N'Supplementary Field 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1417, 110, N'Supplementary Field 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1418, 110, N'Supplementary Field 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1419, 111, N'Code Table ID', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1420, 111, N'Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1421, 111, N'Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1422, 111, N'Short Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1423, 111, N'Email Address', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1424, 111, N'Supplementary Field 1a', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1425, 111, N'Supplementary Field 1b', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1426, 111, N'Supplementary Field 1c', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1427, 111, N'Supplementary Field 1d', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1428, 111, N'Supplementary Field 1e', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1429, 111, N'Supplementary Field 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1430, 111, N'Supplementary Field 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1431, 111, N'Supplementary Field 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1432, 111, N'Supplementary Field 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1433, 111, N'Supplementary Field 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1434, 112, N'Code Table ID', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1435, 112, N'Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1436, 112, N'Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1437, 112, N'Short Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1438, 112, N'Email Address', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1439, 112, N'Supplementary Field 1a', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1440, 112, N'Supplementary Field 1b', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1441, 112, N'Supplementary Field 1c', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1442, 112, N'Supplementary Field 1d', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1443, 112, N'Supplementary Field 1e', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1444, 112, N'Supplementary Field 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1445, 112, N'Supplementary Field 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1446, 112, N'Supplementary Field 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1447, 112, N'Supplementary Field 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1448, 112, N'Supplementary Field 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1449, 113, N'Code Table ID', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1450, 113, N'Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1451, 113, N'Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1452, 113, N'Short Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1453, 113, N'Email Address', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1454, 113, N'Supplementary Field 1a', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1455, 113, N'Supplementary Field 1b', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1456, 113, N'Supplementary Field 1c', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1457, 113, N'Supplementary Field 1d', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1458, 113, N'Supplementary Field 1e', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1459, 113, N'Supplementary Field 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1460, 113, N'Supplementary Field 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1461, 113, N'Supplementary Field 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1462, 113, N'Supplementary Field 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1463, 113, N'Supplementary Field 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1464, 114, N'Code Table ID', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1465, 114, N'Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1466, 114, N'Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1467, 114, N'Short Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1468, 114, N'Email Address', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1469, 114, N'Supplementary Field 1a', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1470, 114, N'Supplementary Field 1b', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1471, 114, N'Supplementary Field 1c', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1472, 114, N'Supplementary Field 1d', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1473, 114, N'Supplementary Field 1e', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1474, 114, N'Supplementary Field 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1475, 114, N'Supplementary Field 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1476, 114, N'Supplementary Field 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1477, 114, N'Supplementary Field 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1478, 114, N'Supplementary Field 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1479, 115, N'Code Table ID', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1480, 115, N'Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1481, 115, N'Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1482, 115, N'Short Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1483, 115, N'Email Address', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1484, 115, N'Supplementary Field 1a', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1485, 115, N'Supplementary Field 1b', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1486, 115, N'Supplementary Field 1c', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1487, 115, N'Supplementary Field 1d', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1488, 115, N'Supplementary Field 1e', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1489, 115, N'Supplementary Field 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1490, 115, N'Supplementary Field 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1491, 115, N'Supplementary Field 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1492, 115, N'Supplementary Field 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1493, 115, N'Supplementary Field 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1494, 116, N'Code Table ID', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1495, 116, N'Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1496, 116, N'Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1497, 116, N'Short Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1498, 116, N'Email Address', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1499, 116, N'Supplementary Field 1a', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1500, 116, N'Supplementary Field 1b', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1501, 116, N'Supplementary Field 1c', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1502, 116, N'Supplementary Field 1d', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1503, 116, N'Supplementary Field 1e', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1504, 116, N'Supplementary Field 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1505, 116, N'Supplementary Field 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1506, 116, N'Supplementary Field 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1507, 116, N'Supplementary Field 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1508, 116, N'Supplementary Field 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1509, 116, N'OSP Indicator', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1510, 116, N'SSP Indicator', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1511, 116, N'Days/Hours', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1512, 117, N'Code Table ID', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1513, 117, N'Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1514, 117, N'Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1515, 117, N'Short Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1516, 117, N'Email Address', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1517, 117, N'Supplementary Field 1a', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1518, 117, N'Supplementary Field 1b', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1519, 117, N'Supplementary Field 1c', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1520, 117, N'Supplementary Field 1d', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1521, 117, N'Supplementary Field 1e', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1522, 117, N'Supplementary Field 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1523, 117, N'Supplementary Field 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1524, 117, N'Supplementary Field 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1525, 117, N'Supplementary Field 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1526, 117, N'Supplementary Field 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1527, 117, N'Absence Type', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1528, 118, N'Sort Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1529, 118, N'Bank Name', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1530, 118, N'Branch', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1531, 118, N'Address 1', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1532, 118, N'Address 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1533, 118, N'Address 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1534, 118, N'Address 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1535, 118, N'Address 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1536, 131, N'Code Table ID', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1537, 131, N'Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1538, 131, N'Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1539, 131, N'Short Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1540, 131, N'Email Address', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1541, 131, N'Supplementary Field 1a', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1542, 131, N'Supplementary Field 1b', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1543, 131, N'Supplementary Field 1c', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1544, 131, N'Supplementary Field 1d', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1545, 131, N'Supplementary Field 1e', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1546, 131, N'Supplementary Field 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1547, 131, N'Supplementary Field 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1548, 131, N'Supplementary Field 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1549, 131, N'Supplementary Field 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1550, 131, N'Supplementary Field 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1551, 132, N'Code Table ID', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1552, 132, N'Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1553, 132, N'Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1554, 132, N'Short Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1555, 132, N'Email Address', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1556, 132, N'Supplementary Field 1a', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1557, 132, N'Supplementary Field 1b', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1558, 132, N'Supplementary Field 1c', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1559, 132, N'Supplementary Field 1d', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1560, 132, N'Supplementary Field 1e', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1561, 132, N'Supplementary Field 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1562, 132, N'Supplementary Field 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1563, 132, N'Supplementary Field 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1564, 132, N'Supplementary Field 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1565, 132, N'Supplementary Field 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1566, 133, N'Code Table ID', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1567, 133, N'Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1568, 133, N'Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1569, 133, N'Short Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1570, 133, N'Email Address', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1571, 133, N'Supplementary Field 1a', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1572, 133, N'Supplementary Field 1b', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1573, 133, N'Supplementary Field 1c', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1574, 133, N'Supplementary Field 1d', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1575, 133, N'Supplementary Field 1e', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1576, 133, N'Supplementary Field 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1577, 133, N'Supplementary Field 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1578, 133, N'Supplementary Field 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1579, 133, N'Supplementary Field 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1580, 133, N'Supplementary Field 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1581, 134, N'Code Table ID', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1582, 134, N'Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1583, 134, N'Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1584, 134, N'Short Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1585, 134, N'Email Address', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1586, 134, N'Supplementary Field 1a', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1587, 134, N'Supplementary Field 1b', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1588, 134, N'Supplementary Field 1c', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1589, 134, N'Supplementary Field 1d', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1590, 134, N'Supplementary Field 1e', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1591, 134, N'Supplementary Field 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1592, 134, N'Supplementary Field 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1593, 134, N'Supplementary Field 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1594, 134, N'Supplementary Field 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1595, 134, N'Supplementary Field 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1596, 135, N'Code Table ID', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1597, 135, N'Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1598, 135, N'Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1599, 135, N'Short Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1600, 135, N'Email Address', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1601, 135, N'Supplementary Field 1a', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1602, 135, N'Supplementary Field 1b', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1603, 135, N'Supplementary Field 1c', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1604, 135, N'Supplementary Field 1d', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1605, 135, N'Supplementary Field 1e', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1606, 135, N'Supplementary Field 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1607, 135, N'Supplementary Field 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1608, 135, N'Supplementary Field 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1609, 135, N'Supplementary Field 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1610, 135, N'Supplementary Field 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1611, 136, N'Code Table ID', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1612, 136, N'Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1613, 136, N'Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1614, 136, N'Short Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1615, 136, N'Email Address', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1616, 136, N'Supplementary Field 1a', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1617, 136, N'Supplementary Field 1b', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1618, 136, N'Supplementary Field 1c', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1619, 136, N'Supplementary Field 1d', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1620, 136, N'Supplementary Field 1e', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1621, 136, N'Supplementary Field 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1622, 136, N'Supplementary Field 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1623, 136, N'Supplementary Field 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1624, 136, N'Supplementary Field 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1625, 136, N'Supplementary Field 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1626, 137, N'Code Table ID', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1627, 137, N'Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1628, 137, N'Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1629, 137, N'Short Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1630, 137, N'Email Address', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1631, 137, N'Supplementary Field 1a', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1632, 137, N'Supplementary Field 1b', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1633, 137, N'Supplementary Field 1c', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1634, 137, N'Supplementary Field 1d', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1635, 137, N'Supplementary Field 1e', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1636, 137, N'Supplementary Field 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1637, 137, N'Supplementary Field 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1638, 137, N'Supplementary Field 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1639, 137, N'Supplementary Field 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1640, 137, N'Supplementary Field 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1641, 138, N'Code Table ID', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1642, 138, N'Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1643, 138, N'Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1644, 138, N'Short Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1645, 138, N'Email Address', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1646, 138, N'Supplementary Field 1a', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1647, 138, N'Supplementary Field 1b', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1648, 138, N'Supplementary Field 1c', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1649, 138, N'Supplementary Field 1d', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1650, 138, N'Supplementary Field 1e', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1651, 138, N'Supplementary Field 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1652, 138, N'Supplementary Field 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1653, 138, N'Supplementary Field 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1654, 138, N'Supplementary Field 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1655, 138, N'Supplementary Field 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1656, 139, N'Code Table ID', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1657, 139, N'Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1658, 139, N'Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1659, 139, N'Short Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1660, 139, N'Email Address', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1661, 139, N'Supplementary Field 1a', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1662, 139, N'Supplementary Field 1b', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1663, 139, N'Supplementary Field 1c', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1664, 139, N'Supplementary Field 1d', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1665, 139, N'Supplementary Field 1e', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1666, 139, N'Supplementary Field 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1667, 139, N'Supplementary Field 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1668, 139, N'Supplementary Field 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1669, 139, N'Supplementary Field 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1670, 139, N'Supplementary Field 6', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1671, 140, N'Code Table ID', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1672, 140, N'Code', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1673, 140, N'Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1674, 140, N'Short Description', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1675, 140, N'Email Address', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1676, 140, N'Supplementary Field 1a', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1677, 140, N'Supplementary Field 1b', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1678, 140, N'Supplementary Field 1c', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1679, 140, N'Supplementary Field 1d', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1680, 140, N'Supplementary Field 1e', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1681, 140, N'Supplementary Field 2', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1682, 140, N'Supplementary Field 3', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1683, 140, N'Supplementary Field 4', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1684, 140, N'Supplementary Field 5', NULL, -1, NULL, NULL, NULL, 0)
GO
INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [ColumnID], [Lookup]) VALUES (1685, 140, N'Supplementary Field 6', NULL, -1, NULL, NULL, NULL, 0)
GO


*/