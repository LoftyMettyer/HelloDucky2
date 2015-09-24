
CREATE TABLE [dbo].[UserDefined1](
	[Id] [int] IDENTITY(1,1) NOT NULL,
	[column1] [varchar](50) NULL,
	[column2] [varchar](255) NULL,
	[column3] [nvarchar](255) NULL,
	[column4] [nvarchar](255) NULL,
	[column5] [datetime2](7) NULL,
	[column6] [nvarchar](MAX) NULL,
	[column7] [bit] NULL,
	[column8] [nvarchar](50) NULL,
	[column9] [nvarchar](50) NULL,
	[column10] numeric(6,2) NULL,
	[column11] [nvarchar](max) NULL,
	[column12] [nvarchar](max) NULL,
	[column13] numeric(6,2) NULL,
	[column14] numeric(6,2) NULL,
	[column15] numeric(6,2) NULL,
	[column16] numeric(6,2) NULL,
	[column25] int NULL,
	[column26] int NULL,
 CONSTRAINT [PK_UserDefined1] PRIMARY KEY CLUSTERED 
(
	[id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[UserDefined2](
	[Id] [int] IDENTITY(1,1) NOT NULL,
	[id_UserDefined1] [int] NULL,
	[column17] [datetime2](7) NULL,
	[column18] [datetime2](7) NULL,
	[column19] numeric(6,2) NULL,
	[column20] [varchar](255) NULL,
	[column21] [varchar](255) NULL,
	[column22] [varchar](255) NULL,
	[column23] [varchar](255) NULL,
	[column24] [bit] NULL,
 CONSTRAINT [PK_UserDefined2] PRIMARY KEY CLUSTERED 
(
	[id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO


CREATE TABLE [dbo].[UserDefined4](
	[Id] [int] IDENTITY(1,1) NOT NULL,
	[column34] int NULL,
	[column35] [nvarchar](5) NOT NULL,
	[column36] [nvarchar](max) NULL,
 CONSTRAINT [PK_UserDefined4] PRIMARY KEY CLUSTERED 
(
	[Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[UserDefined5](
	[Id] [int] IDENTITY(1,1) NOT NULL,
	[column37] int NULL,
	[column38] [nvarchar](5) NOT NULL,
	[column39] [nvarchar](max) NULL,
 CONSTRAINT [PK_UserDefined5] PRIMARY KEY CLUSTERED 
(
	[Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

INSERT dynamicTable (Name, Description, Type) values ('Personnel', 'employee records', 0)
INSERT dynamicTable (Name, Description, Type) values ('Holiday Requests', 'I wanna day off table', 0)
INSERT dynamicTable (Name, Description, Type) values ('Absence', 'Actual Absence Days', 0)
INSERT dynamicTable (Name, Description, Type) values ('Title', 'Person title', 1)
INSERT dynamicTable (Name, Description, Type) values ('Absence Type', 'Absence Type', 1)


INSERT DynamicColumn (TableID, Name, DataType) VALUES (1, 'Forenames', 1)
INSERT DynamicColumn (TableID, Name, DataType) VALUES (1, 'Surname', 1)
INSERT DynamicColumn (TableID, Name, DataType) VALUES (1, 'WorkEmail', 1)
INSERT DynamicColumn (TableID, Name, DataType) VALUES (1, 'Password', 1)
INSERT DynamicColumn (TableID, Name, DataType) VALUES (1, 'BirthDate', 2)
INSERT DynamicColumn (TableID, Name, DataType) VALUES (1, 'AdditionalComments', 1)
INSERT DynamicColumn (TableID, Name, DataType) VALUES (1, 'Checkbox1', 3)
INSERT DynamicColumn (TableID, Name, DataType) VALUES (1, 'IhaveASecret', 1);
INSERT DynamicColumn (TableID, Name, DataType) VALUES (1, 'Gender', 1);
INSERT DynamicColumn (TableID, Name, DataType) VALUES (1, 'BrowserId', 4);
INSERT DynamicColumn (TableID, Name, DataType) VALUES (1, 'Fullname', 1);
INSERT DynamicColumn (TableID, Name, DataType) VALUES (1, 'AbsenceIn', 1);
INSERT DynamicColumn (TableID, Name, DataType) VALUES (1, 'HolidayEntitlement', 4);
INSERT DynamicColumn (TableID, Name, DataType) VALUES (1, 'HolidayTaken', 4);
INSERT DynamicColumn (TableID, Name, DataType) VALUES (1, 'HolidayBroughtForward', 4);
INSERT DynamicColumn (TableID, Name, DataType) VALUES (1, 'HolidayBalance', 4);

INSERT DynamicColumn (TableID, Name, DataType) VALUES (2, 'StartDate', 2);
INSERT DynamicColumn (TableID, Name, DataType) VALUES (2, 'EndDate', 2);
INSERT DynamicColumn (TableID, Name, DataType) VALUES (2, 'DurationHours', 4);
INSERT DynamicColumn (TableID, Name, DataType) VALUES (2, 'EmployeesNotes', 1);
INSERT DynamicColumn (TableID, Name, DataType) VALUES (2, 'RequestType', 1);
INSERT DynamicColumn (TableID, Name, DataType) VALUES (2, 'StartSession', 1);
INSERT DynamicColumn (TableID, Name, DataType) VALUES (2, 'EndSession', 1);
INSERT DynamicColumn (TableID, Name, DataType) VALUES (2, 'IsApproved', 3);

INSERT DynamicColumn (TableID, Name, DataType, LookupTableId) VALUES (1, 'Title', 6, 4);
INSERT DynamicColumn (TableID, Name, DataType) VALUES (1, 'Age', 5);

INSERT DynamicColumn (TableID, Name, DataType) VALUES (3, 'StartDate', 2);
INSERT DynamicColumn (TableID, Name, DataType) VALUES (3, 'EndDate', 2);
INSERT DynamicColumn (TableID, Name, DataType) VALUES (3, 'Duration', 4);
INSERT DynamicColumn (TableID, Name, DataType) VALUES (3, 'Notes', 1);
INSERT DynamicColumn (TableID, Name, DataType) VALUES (3, 'Type', 1);
INSERT DynamicColumn (TableID, Name, DataType) VALUES (3, 'StartSession', 1);
INSERT DynamicColumn (TableID, Name, DataType) VALUES (3, 'EndSession', 1);

INSERT DynamicColumn (TableID, Name, DataType) VALUES (4, 'KeyId', 1);
INSERT DynamicColumn (TableID, Name, DataType) VALUES (4, 'LanguageId', 1);
INSERT DynamicColumn (TableID, Name, DataType) VALUES (4, 'Value', 5);

INSERT DynamicColumn (TableID, Name, DataType) VALUES (5, 'KeyId', 1);
INSERT DynamicColumn (TableID, Name, DataType) VALUES (5, 'LanguageId', 1);
INSERT DynamicColumn (TableID, Name, DataType) VALUES (5, 'Value', 5);

INSERT DynamicColumn (TableID, Name, DataType) VALUES (2, 'Category', 1);
INSERT DynamicColumn (TableID, Name, DataType) VALUES (2, 'SubCategory', 1);



-- Workflow processes
DELETE FROM [WebFormButton]
DELETE FROM [WebFormFieldOption]
DELETE FROM [WebFormField]
DELETE FROM [WebForm]
DELETE FROM [ProcessElement]
DELETE FROM [Process]

DBCC CHECKIDENT ('WebForm', reseed, 0)
DBCC CHECKIDENT ('WebFormButton', reseed, 0)
DBCC CHECKIDENT ('WebFormFieldOption', reseed, 0)
DBCC CHECKIDENT ('WebFormField', reseed, 0)
DBCC CHECKIDENT ('ProcessElement', reseed, 0)
DBCC CHECKIDENT ('Process', reseed, 0)

INSERT Process (Name) VALUES ('Change Personal Details')
INSERT Process (Name) VALUES ('Holiday Request')

INSERT [WebForm] VALUES ('Edit Personnel Details')
INSERT [WebForm] VALUES ('Enter Holiday details')

INSERT [ProcessElement] (Process_Id, [Type], [Sequence]) VALUES (1, 1, 1);
INSERT [ProcessElement] (Process_Id, [Type], [Sequence]) VALUES (1, 2, 3);
INSERT [ProcessElement] (Process_Id, [Type], [Sequence], [WebForm_Id]) VALUES (1, 3, 2, 1);

INSERT [ProcessElement] (Process_Id, [Type], [Sequence]) VALUES (2, 1, 1);
INSERT [ProcessElement] (Process_Id, [Type], [Sequence]) VALUES (2, 2, 4);
INSERT [ProcessElement] (Process_Id, [Type], [Sequence], [WebForm_Id]) VALUES (2, 3, 2, 1);
INSERT [ProcessElement] (Process_Id, [Type], [Sequence]) VALUES (2, 4, 3);


INSERT [WebFormField] ([sequence], columnid, title, value, required, disabled, WebForm_id, [type])
	values (1, 1, 'First Name', 'John', 1, 0, 1, 'textfield')
INSERT [WebFormField] ([sequence], columnid, title, value, required, disabled, WebForm_id, [type])
	values (2, 2, 'Last Name', 'Smith', 1, 0, 1, 'textfield')
INSERT [WebFormField] ([sequence], columnid, title, value, required, disabled, WebForm_id, [type])
	values (4, 3, 'Work Email Address', 'test@example.com', 1, 0, 1, 'textfield')
INSERT [WebFormField] ([sequence], columnid, title, value, required, disabled, WebForm_id, [type])
	values (5, 4, 'Password', '', 1, 0, 1, 'textfield')
INSERT [WebFormField] ([sequence], columnid, title, value, required, disabled, WebForm_id, [type])
	values (6, 5, 'Birth Date', '17.09.1971', 1, 0, 1, 'date')
INSERT [WebFormField] ([sequence], columnid, title, value, required, disabled, WebForm_id, [type])
	values (8, 6, 'Additional Comments', 'Please type here...', 1, 0, 1, 'textfield')
INSERT [WebFormField] ([sequence], columnid, title, value, required, disabled, WebForm_id, [type])
	values (9, 7, 'I accept the terms and conditions', '0', 1, 0, 1, 'checkbox')
INSERT [WebFormField] ([sequence], columnid, title, value, required, disabled, WebForm_id, [type])
	values (10, 8, 'I have a secret', '', 1, 0, 1, 'textfield')
INSERT [WebFormField] ([sequence], columnid, title, value, required, disabled, WebForm_id, [type])
	values (3, 9, 'Gender', '2', 1, 0, 1, 'dropdown')
INSERT [WebFormField] ([sequence], columnid, title, value, required, disabled, WebForm_id, [type])
	values (7, 10, 'Browser', '2', 1, 0, 1, 'dropdown')
INSERT [WebFormField] ([sequence], columnid, title, value, required, disabled, WebForm_id, [type])
	values (2, 11, 'Full Name', 'from db', 1, 0, 1, 'textfield')
INSERT [WebFormField] ([sequence], columnid, title, value, required, disabled, WebForm_id, [type])
	values (3, 12, 'Absence In', 'Days', 1, 0, 1, 'textfield')
INSERT [WebFormField] ([sequence], columnid, title, value, required, disabled, WebForm_id, [type])
	values (4, 13, 'holsiday entitlement', '', 1, 0, 1, 'textfield')
INSERT [WebFormField] ([sequence], columnid, title, value, required, disabled, WebForm_id, [type])
	values (5, 14, 'holidays taken', '', 1, 0, 1, 'textfield')
INSERT [WebFormField] ([sequence], columnid, title, value, required, disabled, WebForm_id, [type])
	values (6, 15, 'balance brought forward', '', 1, 0, 1, 'textfield')
INSERT [WebFormField] ([sequence], columnid, title, value, required, disabled, WebForm_id, [type])
	values (7, 16, 'Holiday Balance', '', 1, 0, 1, 'textfield')
INSERT [WebFormField] ([sequence], columnid, title, value, required, disabled, WebForm_id, [type])
	values (11, 25, 'title', '', 1, 0, 1, 'dropdown')


INSERT [WebFormField] ([sequence], columnid, title, value, required, disabled, WebForm_id, [type])
	values (9, 17, 'Start Date', '', 1, 0, 2, 'date')
INSERT [WebFormField] ([sequence], columnid, title, value, required, disabled, WebForm_id, [type])
	values (11, 18, 'End Date', '', 1, 0, 2, 'date')
INSERT [WebFormField] ([sequence], columnid, title, value, required, disabled, WebForm_id, [type])
--	values (13, 19, 'Duration Hours', '', 1, 0, 16, 'textfield')
	values (13, 19, 'Holiday Summary', '5,30,7.5,22', 1, 0, 2, 'bulletgraph')

INSERT [WebFormField] ([sequence], columnid, title, value, required, disabled, WebForm_id, [type])
	values (14, 20, 'Employees Notes', '', 0, 0, 2, 'textarea')
INSERT [WebFormField] ([sequence], columnid, title, value, required, disabled, WebForm_id, [type])
	values (8, 21, 'Request Type', '', 1, 0, 2, 'dropdown')
INSERT [WebFormField] ([sequence], columnid, title, value, required, disabled, WebForm_id, [type])
	values (10, 22, 'Start Session', 'AM', 1, 0, 2, 'dropdown')
INSERT [WebFormField] ([sequence], columnid, title, value, required, disabled, WebForm_id, [type])
	values (12, 23, 'End Session', 'PM', 0, 0, 2, 'dropdown')
INSERT [WebFormField] ([sequence], columnid, title, value, required, disabled, WebForm_id, [type])
	values (13, 24, 'Is Approved', 0, 0, 0, 2, 'checkbox')




INSERT [WebFormFieldOption] ([title], [value], [WebFormField_id])
	VALUES ('Girl', 2, 9)
INSERT [WebFormFieldOption] ([title], [value], [WebFormField_id])
	VALUES ('Bloke', 10, 9)

INSERT [WebFormFieldOption] ([title], [value], [WebFormField_id])
	VALUES ('Holiday', 1, 21)
INSERT [WebFormFieldOption] ([title], [value], [WebFormField_id])
	VALUES ('Jury Service', 2, 21)
INSERT [WebFormFieldOption] ([title], [value], [WebFormField_id])
	VALUES ('Sick', 3, 21)
INSERT [WebFormFieldOption] ([title], [value], [WebFormField_id])
	VALUES ('Compassionate Leave', 4, 21)
INSERT [WebFormFieldOption] ([title], [value], [WebFormField_id])
	VALUES ('Duvert Day', 5, 21)

INSERT [WebFormFieldOption] ([title], [value], [WebFormField_id])
	VALUES ('AM', 1, 22)
INSERT [WebFormFieldOption] ([title], [value], [WebFormField_id])
	VALUES ('PM', 2, 22)

INSERT [WebFormFieldOption] ([title], [value], [WebFormField_id])
	VALUES ('AM', 1, 23)
INSERT [WebFormFieldOption] ([title], [value], [WebFormField_id])
	VALUES ('PM', 2, 23)

INSERT WebFormButton(title, targeturl, WebForm_id, [action])
	VALUES ('Save', '/api/data/SubmitStep/{0}', 1, 1)
INSERT WebFormButton(title, targeturl, WebForm_id, [action])
	VALUES ('Cancel', '', 1, 3)
INSERT WebFormButton(title, targeturl, WebForm_id, [action])
	VALUES ('Request', '/api/resource/saveprocess/step/{0}', 2, 1)
INSERT WebFormButton(title, targeturl, WebForm_id, [action])
	VALUES ('Save for later', '/api/resource/saveprocess/saveforlater/{0}', 2, 2)
INSERT WebFormButton(title, targeturl, WebForm_id, [action])
	VALUES ('Forget it', '', 2, 3)
INSERT WebFormButton(title, targeturl, WebForm_id, [action])
	VALUES ('Refresh Me', '/api/resource/saveprocess/refreshStep/{0}', 2, 4)



--SELECT * FROM DynamicAttribute
select * from dynamicTable
select * from DynamicColumn
select * from WebForm
select * from [dbo].[WebFormField]
select * from [UserDefined1]

--INSERT [User] (userid, recordid) values ('088C6A78-E14A-41B0-AD93-4FB7D3ADE96C', 1)
--insert [dbo].[UserDefined1] ([Column1], [Column2], [Column9], [Column5], [Column7]) values ('fisrt', 'surn', 'M', getdate(), 1)



--DELETE FROM [PersonnelPermission]
--INSERT [PersonnelPermission] (ID_1, gROUPid) SELECT ID_1+1, gROUPid FROM [Nexus_OldThurs].[dbo].[PersonnelPermission]

/*

CREATE TABLE [dbo].[User](
	[UserId] [uniqueidentifier] NULL,
	[Email] [nvarchar](255) NULL,
	[RecordId] [int] NULL
) ON [PRIMARY]

GO

CREATE TABLE [dbo].[Group](
	[Id] [int] IDENTITY(1,1) NOT NULL,
	[Name] [nvarchar](255) NULL,
	[Description] [nvarchar](max) NULL,
 CONSTRAINT [PK_dbo.Group] PRIMARY KEY CLUSTERED 
(
	[Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO

CREATE TABLE [dbo].[GroupPermission](
	[Id] [int] IDENTITY(1,1) NOT NULL,
	[GroupId] [int] NOT NULL,
	[PermissionId] [int] NOT NULL,
	[IsDeny] [bit] NOT NULL,
 CONSTRAINT [PK_GroupPermission] PRIMARY KEY CLUSTERED 
(
	[Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

CREATE TABLE [dbo].[PermissionCategory](
	[Id] [int] IDENTITY(1,1) NOT NULL,
	[KeyName] [nvarchar](25) NULL,
	[Description] [nvarchar](255) NULL,
 CONSTRAINT [PK_dbo.PermissionCategory] PRIMARY KEY CLUSTERED 
(
	[Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

CREATE TABLE [dbo].[PermissionFacet](
	[Id] [int] IDENTITY(1,1) NOT NULL,
	[Name] [nvarchar](25) NULL,
 CONSTRAINT [PK_dbo.PermissionFacet] PRIMARY KEY CLUSTERED 
(
	[Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

CREATE TABLE [dbo].[PermissionItem](
	[Id] [int] IDENTITY(1,1) NOT NULL,
	[PermissionCategoryId] [int] NOT NULL,
	[PermissionFacetId] [int] NOT NULL,
 CONSTRAINT [PK_PermissionItem] PRIMARY KEY CLUSTERED 
(
	[Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

CREATE TABLE [dbo].[PersonnelPermission](
	[id_1] [int] NOT NULL,
	[id] [int] IDENTITY(1,1) NOT NULL,
	[GroupId] [int] NOT NULL,
 CONSTRAINT [PK_Personnel_Roles] PRIMARY KEY CLUSTERED 
(
	[id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

INSERT dbo.UserDefined5 (column37, column38, column39) VALUES (1,'en-gb','Holiday')
INSERT dbo.UserDefined5 (column37, column38, column39) VALUES (1,'fr-fr','Vacances')
INSERT dbo.UserDefined5 (column37, column38, column39) VALUES (2,'en-gb','Sick')
INSERT dbo.UserDefined5 (column37, column38, column39) VALUES (2,'fr-fr','Malade')
INSERT dbo.UserDefined5 (column37, column38, column39) VALUES (3,'en-gb','Jury Service')
INSERT dbo.UserDefined5 (column37, column38, column39) VALUES (3,'fr-fr','Service Jury')
INSERT dbo.UserDefined5 (column37, column38, column39) VALUES (4,'en-gb','Duvet Day')
INSERT dbo.UserDefined5 (column37, column38, column39) VALUES (4,'fr-fr','Jour Couette')
INSERT dbo.UserDefined5 (column37, column38, column39) VALUES (5,'en-gb','Compassionate Leave')
INSERT dbo.UserDefined5 (column37, column38, column39) VALUES (5,'fr-fr','Congé Pour Raisons Familiales')

INSERT dbo.Process (Name) VALUES ('Holiday Request');

INSERT [User] (UserId, RecordID) SELECT UserId, RecordID FROM [Nexus_OldThurs].[dbo].[User]
INSERT [Group] (Name, Description) SELECT Name, Description FROM [Nexus_OldThurs].[dbo].[Group]
INSERT [GroupPermission] (GroupId, PermissionId, IsDeny) SELECT GroupId, PermissionId, IsDeny FROM [Nexus_OldThurs].[dbo].[GroupPermission]
INSERT [PermissionCategory] (KeyName, Description) SELECT KeyName, Description FROM [Nexus_OldThurs].[dbo].[PermissionCategory]
INSERT [PersonnelPermission] (id_1, GroupID) SELECT id_1, GroupId FROM [Nexus_OldThurs].[dbo].[PersonnelPermission]
INSERT [PermissionFacet] (Name) SELECT Name FROM [Nexus_OldThurs].[dbo].[PermissionFacet]
INSERT [PermissionItem] (PermissionCategoryId, PermissionFacetId) SELECT PermissionCategoryId, PermissionFacetId FROM [Nexus_OldThurs].[dbo].[PermissionItem]

INSERT INTO [UserDefined1] (column1,column2,column26,column3,column4,column9,column5,column6,column7,column8,column10,column11,column12,column13,column14,column15,column16)
	select [Surname],[Forenames],[TitleId],[WorkEmail],[Password],[Gender],[BirthDate],[AdditionalComments],[Checkbox1],[iHaveASecret],[BrowserID],[FullName],[AbsenceIn],[HolidayEntitlement],[HolidayTaken],[HolidayBroughtForward],[HolidayBalance] from [Nexus_OldThurs].[dbo].[Personnel]

INSERT INTO [UserDefined4] (column34, column35, column36) SELECT KeyId, LanguageId, Value FROM [Nexus_OldThurs].[dbo].[Title]


GO
CREATE PROCEDURE [dbo].[GetWelcomeMessageData](
	@UserId uniqueidentifier,
	@Language nvarchar(50))
AS
BEGIN

	-- This needs to be a cleaner check! (temporary for demo purposes)
	IF NOT EXISTS(SELECT [column35] FROM [UserDefined4] WHERE [column35] = @Language)
		SET @Language = 'EN-GB';

	SELECT @UserID AS [UserId], @Language AS [Language], t.column36 + ' ' + p.column2 + ' ' + p.column1 AS [Message], GETDATE()-1 AS [LastLoggedOn]
		FROM [User] u
		INNER JOIN [UserDefined1] p ON p.Id = u.RecordID
		INNER JOIN [UserDefined4] t ON t.column35 = @Language AND t.column34 = p.column26
	WHERE u.UserID = @UserId;

END
GO

CREATE PROCEDURE [dbo].[RegisterNewUser](
	@email nvarchar(255),
	@userId nvarchar(255)
	)
AS
BEGIN

	DECLARE @recordID int = 0,
			@Status integer = 0;

	IF EXISTS(SELECT * FROM [User] WHERE Email = @email) SET @Status = 1
	ELSE
	BEGIN
	
		SELECT @recordID = Id FROM UserDefined1 WHERE column3 = @email;

		IF ISNULL(@recordID, 0) = 0 SET @Status = 2;
		ELSE
		BEGIN
			SET @Status = 3;
 
			INSERT [User] (UserID, Email, RecordID)
				VALUES (@userId, @email, @recordID);
		END
	END

	SELECT @Status AS [Status];

END

GO

CREATE VIEW [dbo].[UserRole]
AS
SELECT pp.id, u.UserId, c.KeyName + f.Name AS [Name]
	FROM dbo.PersonnelPermission AS pp 
		INNER JOIN dbo.[User] u ON u.RecordId = pp.id_1
		INNER JOIN dbo.[Group] g ON g.Id = pp.GroupId
		INNER JOIN dbo.[GroupPermission] gp ON gp.GroupId = g.Id
		INNER JOIN dbo.[PermissionItem] p ON p.Id = gp.PermissionId
		INNER JOIN dbo.[PermissionCategory] c ON p.PermissionCategoryId = c.Id
		INNER JOIN dbo.[PermissionFacet] f ON f.Id = p.PermissionFacetId

GO


CREATE VIEW dbo.[Lookups]
AS
	SELECT id, column36 AS [title], column34 AS value, 25 AS WebFormField_id, column35 AS [Language] FROM userdefined4
	UNION
	SELECT id, column39 AS [title], column37 AS value, 21 AS WebFormField_id, column38 AS [Language] FROM userdefined5

GO



*/


--exec [dbo].[GetWelcomeMessageData] '7A9043B6-AE62-43F1-B441-95173C45BFE1', 'fr-fr'
--exec [dbo].[GetWelcomeMessageData] '7A9043B6-AE62-43F1-B441-95173C45BFE1', 'en-GB'

--select * from Lookups



select * from process