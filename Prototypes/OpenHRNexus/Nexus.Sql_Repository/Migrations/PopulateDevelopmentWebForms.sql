DELETE FROM [WebFormFieldOption]
DELETE FROM [WebFormField]
DELETE FROM [WebForm]


DECLARE @insertedForm table (id integer);
DECLARE @webFormId integer,
		@WebFormFieldId integer;

-- Basic tables
INSERT [DynamicTable] (name, Type) VALUES ('Personnel', 0)
INSERT [DynamicTable] (name, Type) VALUES ('Absence', 0)


-- Original dummy personnel screen
INSERT dbo.[WebForm] (Name) OUTPUT inserted.id INTO @insertedForm VALUES ('Personnel');

SELECT @webFormId = id FROM @insertedForm;

INSERT INTO dbo.[WebFormField] ([WebForm_Id], [field_id], [field_columnId]  ,[field_title], [field_type], [field_value], [field_required], [field_disabled]) VALUES
    (@webFormId, 1, 1, 'First Name', 'textfield', 'John', 1, 0),
    (@webFormId, 2, 2, 'Last Name', 'textfield', 'Smith', 1, 0),
    (@webFormId, 4, 3, 'Email Address', 'email', 'test@example.com', 1, 0),
    (@webFormId, 5, 4, 'Password', 'password', '', 1, 0),
    (@webFormId, 6, 5, 'Birth Date', 'date', '17.09.1971', 1, 0),
    (@webFormId, 8, 6, 'Additional Comments', 'textarea', 'Please type here...', 0, 0),
    (@webFormId, 9, 7, 'I accept the terms and conditions', 'checkbox', '0', 1, 0),
    (@webFormId, 10, 8, 'I have a secret', 'hidden', 'X', 0, 0);


DELETE FROM @insertedForm;
INSERT INTO dbo.[WebFormField] ([WebForm_Id], [field_id], [field_columnId]  ,[field_title], [field_type], [field_value], [field_required], [field_disabled])
	OUTPUT inserted.id INTO @insertedForm VALUES
	(@webFormId, 3, 9, 'Gender', 'radio', '2', 1, 0);

SELECT @WebFormFieldId = ID FROM @insertedForm;
INSERT INTO dbo.[WebFormFieldOption] ([option_title], [option_value], [WebFormField_id]) VALUES
           ('Female', 2, @WebFormFieldId),
           ('Male', 10, @WebFormFieldId);


DELETE FROM @insertedForm;
INSERT INTO dbo.[WebFormField] ([WebForm_Id], [field_id], [field_columnId]  ,[field_title], [field_type], [field_value], [field_required], [field_disabled])
	OUTPUT inserted.id INTO @insertedForm VALUES
	(@webFormId, 7, 1, 'Your browser', 'dropdown', '2', 0, 0);

SELECT @WebFormFieldId = ID FROM @insertedForm;
INSERT INTO dbo.[WebFormFieldOption] ([option_title], [option_value], [WebFormField_id]) VALUES
           ('--Please Select--', 1, @WebFormFieldId),
           ('Internet Explorer', 2, @WebFormFieldId),
           ('Google Chrome', 3, @WebFormFieldId),
           ('Mozilla Firefox', 4, @WebFormFieldId);



-- Dummy Absence screen
DELETE FROM @insertedForm;
INSERT dbo.[WebForm] (Name) OUTPUT inserted.id INTO @insertedForm VALUES ('Absence Request');

SELECT @webFormId = id FROM @insertedForm;

INSERT INTO dbo.[WebFormField] ([WebForm_Id], [field_id], [field_columnId]  ,[field_title], [field_type], [field_value], [field_required], [field_disabled]) VALUES
    (@webFormId, 2, 22, 'Full Name', 'textfield', 'from database', 0, 1),
    (@webFormId, 3, 23, 'Absence In', 'email', 'Days', 0, 1),
    (@webFormId, 4, 24, 'Holiday Entitlement', 'textfield', 'from database', 0, 1),
    (@webFormId, 5, 25, 'Holiday Taken', 'textfield', 'from database', 0, 1),
    (@webFormId, 6, 26, 'Holiday Brought Forward', 'textfield', 'from database...', 0, 1),
    (@webFormId, 7, 27, 'Holiday Balance', 'textfield', 'db goes here', 0, 1),
    (@webFormId, 9, 28, 'Start Date', 'date', '', 1, 0),
    (@webFormId, 11, 29, 'End Date', 'date', '', 1, 0),
    (@webFormId, 13, 30, 'Duration Hours', 'textfield', '', 0, 0),
    (@webFormId, 14, 31, 'Employees Notes', 'textfield', '', 0, 0);


DELETE FROM @insertedForm;
INSERT INTO dbo.[WebFormField] ([WebForm_Id], [field_id], [field_columnId]  ,[field_title], [field_type], [field_value], [field_required], [field_disabled])
	OUTPUT inserted.id INTO @insertedForm VALUES
    (@webFormId, 8, 28, 'Request Type', 'dropdown', '', 1, 0);

SELECT @WebFormFieldId = ID FROM @insertedForm;
INSERT INTO dbo.[WebFormFieldOption] ([option_title], [option_value], [WebFormField_id]) VALUES
           ('Holiday', 1, @WebFormFieldId),
           ('Jury Service', 2, @WebFormFieldId),
           ('Duvet Day', 3, @WebFormFieldId);

DELETE FROM @insertedForm;
INSERT INTO dbo.[WebFormField] ([WebForm_Id], [field_id], [field_columnId]  ,[field_title], [field_type], [field_value], [field_required], [field_disabled])
	OUTPUT inserted.id INTO @insertedForm VALUES
    (@webFormId, 10, 28, 'Start Session', 'dropdown', '', 1, 0);

SELECT @WebFormFieldId = ID FROM @insertedForm;
INSERT INTO dbo.[WebFormFieldOption] ([option_title], [option_value], [WebFormField_id]) VALUES
           ('AM', 1, @WebFormFieldId),
           ('PM', 2, @WebFormFieldId);


DELETE FROM @insertedForm;
INSERT INTO dbo.[WebFormField] ([WebForm_Id], [field_id], [field_columnId]  ,[field_title], [field_type], [field_value], [field_required], [field_disabled])
	OUTPUT inserted.id INTO @insertedForm VALUES
    (@webFormId, 12, 28, 'End Session', 'dropdown', '', 1, 0);

SELECT @WebFormFieldId = ID FROM @insertedForm;
INSERT INTO dbo.[WebFormFieldOption] ([option_title], [option_value], [WebFormField_id]) VALUES
           ('AM', 1, @WebFormFieldId),
           ('PM', 2, @WebFormFieldId);



--GO

select * from [WebForm]
select * from [WebFormField]
select * from [WebFormFieldOption]
