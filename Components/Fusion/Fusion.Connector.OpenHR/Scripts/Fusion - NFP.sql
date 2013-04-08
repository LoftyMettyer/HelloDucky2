

IF EXISTS (SELECT * FROM sys.views WHERE object_id = object_ID(N'[fusion].[staffContract]'))
	DROP VIEW [fusion].[staffContract];

IF EXISTS (SELECT * FROM sys.views WHERE object_id = object_ID(N'[fusion].[staffSkill]'))
	DROP VIEW [fusion].[staffSkill];

IF EXISTS (SELECT * FROM sys.views WHERE object_id = object_ID(N'[fusion].[staffLegalDocument]'))
	DROP VIEW [fusion].[staffLegalDocument];

IF EXISTS (SELECT * FROM sys.views WHERE object_id = object_ID(N'[fusion].[staff]'))
	DROP VIEW [fusion].[staff];

IF EXISTS (SELECT * FROM sys.views WHERE object_id = object_ID(N'[fusion].[staffContact]'))
	DROP VIEW [fusion].[staffContact];

IF EXISTS (SELECT * FROM sys.views WHERE object_id = object_ID(N'[fusion].[staffTimesheet]'))
	DROP VIEW [fusion].[staffTimesheet];

GO

-- Fusion views of the OpenHR database
-- This is the bit that needs customising. Do not alter the AS clauses

CREATE VIEW [fusion].[staff]
AS
SELECT ID as [StaffID]
	, title AS [title]
	, forenames AS [forenames]
	, surname AS [surname]
	, Known_As AS [preferredName]
	, staff_number AS [payrollNumber]
	, date_of_birth AS [DOB]
	, employee_type AS [employeeType]
	, work_mobile AS [workMobile]
	, personal_mobile AS [personalMobile]
	, work_telephone AS [workPhoneNumber]
	, home_telephone AS [homePhoneNumber]
	, email_work AS [email]
	, email_home AS [personalEmail]
	, gender AS [gender]
	, start_date AS [startDate]
	, leaving_date AS [leavingDate]
	, leaving_reason AS [leavingReason]
	, Division AS [companyName]
	, Post_Title AS [jobTitle]
	, Manager_Ref AS [managerRef]
	, address_1 AS [addressLine1]
	, address_2 AS [addressLine2]
	, address_3 AS [addressLine3]
	, town AS [addressLine4]
	, county AS [addressLine5]
	, postcode AS [postCode]
	, ni_number AS [nationalInsuranceNumber]
	, photograph AS [picture]

FROM dbo.Personnel_Records
GO

CREATE VIEW [fusion].[staffContract]
AS
	SELECT ID						AS [ID_Contract]
		, ID_1						AS [ID_Staff]
		, duty_type					AS [contractName]
		, department				AS [department]
		, location					AS [primarySite]
		, actual_hours				AS [contractedHoursPerWeek]
		, post_hours				AS [maximumHoursPerWeek]
		, appointment_start_date	AS [effectiveFrom]
		, appointment_end_date		AS [effectiveTo]
FROM dbo.Appointments;
GO

CREATE VIEW [fusion].[staffSkill]
AS
	SELECT ID				AS [ID_Skill]
		, ID_1				AS [ID_Staff]
		, course_title		AS [name]
		, start_date		AS [trainingStart]
		, end_date			AS [trainingEnd]
		, valid_from		AS [validFrom]
		, valid_to			AS [validTo]
		, course_code		AS [reference]
		, result			AS [outcome]
		, did_not_attend	AS [didNotAttend]
FROM dbo.Training_Booking;
GO

CREATE VIEW [fusion].[staffLegalDocument]
AS
	SELECT ID					AS [ID_Document]
		, ID_1					AS [ID_Staff]
		, Type					AS [typeName]
		, Valid_From			AS [validFrom]
		, Valid_To				AS [validTo]
		, Document_Reference	AS [documentReference]
		, Secondary_Reference	AS [secondaryReference]
		, Requested_By			AS [requestedBy]
		, Date_Requested		AS [requestedDate]
		, Accepted_By			AS [acceptedBy]
		, Date_Accepted			AS [acceptedDate]
FROM dbo.Legal_Documents;
GO

CREATE VIEW [fusion].[staffContact]
AS
	SELECT ID				AS [ID_Contact]
		, ID_1				AS [ID_Staff]
		, Title				AS [title]
		, Forenames			AS [forenames]
		, Surname			AS [surname]
		, Contact_Type		AS [description]
		, Relationship		AS [relationshipType]
		, Work_Mobile		AS [workMobile]
		, Personal_Mobile	AS [personalMobile]
		, Work_Telephone	AS [workPhoneNumber]
		, Home_Telephone	AS [homePhoneNumber]
		, Email				AS [email]
		, Notes				AS [notes]
		, Address_1			AS [addressLine1]
		, Address_2			AS [addressLine2]
		, Town				AS [addressLine3]
		, County			AS [addressLine4]
		, Country			AS [addressLine5]
		, Postcode			AS [postcode]
	FROM dbo.Contacts

GO

CREATE VIEW [fusion].[staffTimesheet]
AS
	SELECT ID					AS [ID_Timesheet]
		, ID_1					AS [ID_Staff]
		, Timesheet_Date		AS [timesheetDate]
		, Planned_Hours			AS [plannedHours]
		, Worked_Hours			AS [workedHours]
		, TOIL_Hours_Accrued	AS [toilHoursAccrued]
		, Holiday_Hours_Taken	AS [holidayHoursTaken]
		, TOIL_Hours_Taken		AS [toilHoursTaken]
	FROM dbo.Fusion_Timesheet_Submissions

GO


IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[fusion].[pMessageUpdate_StaffChange]') AND xtype = 'P')
	DROP PROCEDURE [fusion].[pMessageUpdate_StaffChange];

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[fusion].[pMessageUpdate_StaffContactChange]') AND xtype = 'P')
	DROP PROCEDURE [fusion].[pMessageUpdate_StaffContactChange];

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[fusion].[pMessageUpdate_StaffContractChange]') AND xtype = 'P')
	DROP PROCEDURE [fusion].[pMessageUpdate_StaffContractChange];

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[fusion].[pMessageUpdate_StaffLegalDocumentChange]') AND xtype = 'P')
	DROP PROCEDURE [fusion].[pMessageUpdate_StaffLegalDocumentChange];

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[fusion].[pMessageUpdate_StaffPictureChange]') AND xtype = 'P')
	DROP PROCEDURE [fusion].[pMessageUpdate_StaffPictureChange];

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[fusion].[pMessageUpdate_StaffSkillChange]') AND xtype = 'P')
	DROP PROCEDURE [fusion].[pMessageUpdate_StaffSkillChange];

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[fusion].[pMessageUpdate_StaffTimesheetSubmission]') AND xtype = 'P')
	DROP PROCEDURE [fusion].[pMessageUpdate_StaffTimesheetSubmission];

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[fusion].[pMessageUpdate_StaffPictureChange]') AND xtype = 'P')
	DROP PROCEDURE [fusion].[pMessageUpdate_StaffPictureChange];

GO

CREATE PROCEDURE fusion.pMessageUpdate_StaffChange(@ID int OUTPUT
	, @title					nvarchar(MAX)			
	, @forenames				nvarchar(MAX)
	, @surname					nvarchar(MAX)
	, @preferredName			nvarchar(MAX)
	, @payrollNumber			nvarchar(MAX)
	, @DOB						datetime
	, @employeeType				nvarchar(MAX)
	, @workMobile				nvarchar(MAX)
	, @personalMobile			nvarchar(MAX)
	, @workPhoneNumber			nvarchar(50)
	, @homePhoneNumber			nvarchar(MAX)
	, @email					nvarchar(MAX)
	, @personalEmail			nvarchar(MAX)
	, @gender					nvarchar(MAX)
	, @startDate				datetime
	, @leavingDate				datetime
	, @leavingReason			nvarchar(MAX)
	, @companyName				nvarchar(MAX)
	, @jobTitle					nvarchar(MAX)
	, @managerRef				nvarchar(MAX)
	, @addressLine1				nvarchar(MAX)
	, @addressLine2				nvarchar(MAX)
	, @addressLine3				nvarchar(MAX)
	, @addressLine4				nvarchar(MAX)
	, @addressLine5				nvarchar(MAX)
	, @postCode					nvarchar(MAX)
	, @nationalInsuranceNumber	nvarchar(MAX)
)
AS
BEGIN

	DECLARE @childID	integer;

	IF ISNULL(@ID,0) = 0
	BEGIN
		INSERT fusion.staff (title, forenames, surname, preferredName, payrollNumber, DOB, employeeType, workMobile,
                            personalMobile, workPhoneNumber, homePhoneNumber, email, personalEmail, gender, startDate, leavingDate,
                            leavingReason, companyName, jobTitle, managerRef, nationalInsuranceNumber)
			VALUES (@title, @forenames, @surname, @preferredName, @payrollNumber, @DOB, @employeeType, @workMobile
					, @personalMobile, @workPhoneNumber, @homePhoneNumber, @email, @personalEmail, @gender, @startDate, @leavingDate
					, @leavingReason, @companyName, @jobTitle, @managerRef, @nationalInsuranceNumber);

		SELECT @ID = MAX(StaffID) FROM fusion.staff;

		INSERT dbo.Address (ID_1, Date_From, Type, Address_1, Address_2, Address_3, Town, County, Postcode)
			VALUES (@ID, GETDATE(), 'Home', @addressLine1, @addressLine2, @addressLine3, @addressLine4, @addressLine5, @postCode);

	END
	ELSE
	BEGIN

		UPDATE fusion.staff SET title = @title, forenames = @forenames, surname = @surname, preferredName = @preferredName, payrollNumber = @payrollNumber
			, DOB = @DOB, employeeType = @employeeType, workMobile = @workMobile
            , personalMobile = @personalMobile, workPhoneNumber = @workPhoneNumber, homePhoneNumber = @homePhoneNumber
			, email = @email, personalEmail = @personalEmail, gender = @gender, startDate = @startDate, leavingDate = @leavingDate
            , leavingReason = @leavingReason, companyName = @companyName, jobTitle = @companyName, managerRef = @managerRef
            , nationalInsuranceNumber = @nationalInsuranceNumber WHERE StaffID = @ID;

		SELECT TOP 1 @childID = ID FROM dbo.Address 
			WHERE ID_1 = @ID AND Date_From <= DATEADD(dd, 0, DATEDIFF(dd, 0, GETDATE())) AND Type = 'Home'
			ORDER BY Date_From DESC;

		UPDATE dbo.Address SET Address_1 = @addressLine1, Address_2 = @addressLine2, Address_3 = @addressLine3
				, Town = @addressLine4, County = @addressLine5, Postcode = @postCode
			WHERE ID = @childID;

	END

END
GO


CREATE PROCEDURE fusion.pMessageUpdate_StaffContractChange(@ID int OUTPUT
	, @staffID					integer
	, @contractName				nvarchar(MAX)
	, @department				nvarchar(MAX)
	, @primarySite				nvarchar(MAX)
	, @contractedHoursPerWeek	numeric(10,2)
	, @maximumHoursPerWeek		numeric(10,2)
	, @effectiveFrom			datetime
	, @effectiveTo				datetime)
AS
BEGIN

	IF ISNULL(@ID,0) = 0
	BEGIN
		INSERT fusion.staffContract(ID_Staff, contractName, department, primarySite, contractedHoursPerWeek
					, maximumHoursPerWeek, effectiveFrom, effectiveTo)
			VALUES (@staffID, @contractName, @department, @primarySite, @contractedHoursPerWeek
					, @maximumHoursPerWeek, @effectiveFrom, @effectiveTo);

		SELECT @ID = MAX(ID_Contract) FROM fusion.staffContract;

	END
	ELSE
	BEGIN

		UPDATE fusion.staffContract SET contractName = @contractName, department = @department, primarySite = @primarySite
			, contractedHoursPerWeek = @contractedHoursPerWeek, maximumHoursPerWeek = @maximumHoursPerWeek
			, effectiveFrom = @effectiveFrom, effectiveTo = @effectiveTo WHERE ID_Contract = @ID

	END

END
GO

CREATE PROCEDURE fusion.pMessageUpdate_StaffContactChange(@ID int OUTPUT
	, @staffID					integer
	, @title					nvarchar(MAX)
	, @forenames				nvarchar(MAX)
	, @surname					nvarchar(MAX)
	, @description				nvarchar(MAX)
	, @relationshipType			nvarchar(MAX)
	, @workMobile				nvarchar(MAX)
	, @personalMobile			nvarchar(MAX)
	, @workPhoneNumber			nvarchar(MAX)
	, @homePhoneNumber			nvarchar(MAX)
	, @email					nvarchar(MAX)
	, @notes					nvarchar(MAX)
	, @addressLine1				nvarchar(MAX)
	, @addressLine2				nvarchar(MAX)
	, @addressLine3				nvarchar(MAX)
	, @addressLine4				nvarchar(MAX)
	, @addressLine5				nvarchar(MAX)
	, @postCode					nvarchar(MAX))
AS
BEGIN

	IF ISNULL(@ID,0) = 0
	BEGIN

		SELECT * FROM fusion.staffContact

		INSERT fusion.staffContact(ID_Staff, title, forenames, surname, [description], relationshipType,
					workMobile, personalMobile, workPhoneNumber, homePhoneNumber, email, notes,
					addressLine1, addressLine2, addressLine3, addressLine4, addressLine5, postcode)
			VALUES (@staffID, @title, @forenames, @surname, @description, @relationshipType,
					@workMobile, @personalMobile, @workPhoneNumber, @homePhoneNumber, @email, @notes,
					@addressLine1, @addressLine2, @addressLine3, @addressLine4, @addressLine5, @postCode);

		SELECT @ID = MAX(ID_Contact) FROM fusion.staffContact;

	END
	ELSE
	BEGIN

		UPDATE fusion.staffContact SET title = @title, forenames = @forenames, surname = @surname, [description] = @description, relationshipType = @relationshipType,
			workMobile = @workMobile, personalMobile = @personalMobile, workPhoneNumber = @workPhoneNumber, homePhoneNumber = @homePhoneNumber,
			email = @email, notes = @notes,
			addressLine1 = @addressLine1, addressLine2 = @addressLine2, addressLine3 = @addressLine3,
			addressLine4 = @addressLine4, addressLine5 = @addressLine5, postcode = @postCode
			WHERE ID_Contact = @ID;

	END

END
GO

CREATE PROCEDURE fusion.pMessageUpdate_StaffLegalDocumentChange(@ID int OUTPUT
	,@staffID				integer
	,@typeName				varchar(MAX)
	,@validFrom				datetime
	,@validTo				datetime
	,@documentReference		varchar(MAX)
	,@secondaryReference	varchar(MAX)
	,@requestedBy			varchar(MAX)
	,@requestedDate			datetime
	,@acceptedBy			varchar(MAX)
	,@acceptedDate			datetime
)
AS
BEGIN

	IF ISNULL(@ID,0) = 0
	BEGIN
		INSERT fusion.staffLegalDocument (ID_Staff, typeName, validFrom, validTo, documentReference, secondaryReference
					, requestedBy, requestedDate, acceptedBy, acceptedDate)
			VALUES (@staffID, @typeName, @validFrom, @validTo, @documentReference, @secondaryReference
					, @requestedBy, @requestedDate, @acceptedBy, @acceptedDate);

		SELECT @ID = MAX(ID_Document) FROM fusion.staffLegalDocument;


	END
	ELSE
	BEGIN

		UPDATE fusion.staffLegalDocument SET typeName = @typeName, validFrom = @validFrom, validTo = @validTo
					, documentReference = @documentReference, secondaryReference = @secondaryReference
					, requestedBy = @requestedBy, requestedDate = @requestedDate
					, acceptedBy = @acceptedBy, acceptedDate = @acceptedDate WHERE ID_Document = @ID;


	END

END
GO

CREATE PROCEDURE fusion.pMessageUpdate_StaffSkillChange(@ID int OUTPUT
	 ,@staffID				integer
	 ,@name					nvarchar(MAX)
	 ,@trainingStart		datetime
	 ,@trainingEnd			datetime
	, @validFrom			datetime
	, @validTo				datetime
	, @reference			nvarchar(MAX)
	, @outcome				nvarchar(MAX)
	, @didNotAttend			bit
)
AS
BEGIN

	DECLARE @childID	integer;

	IF ISNULL(@ID,0) = 0
	BEGIN
		INSERT fusion.staffSkill (ID_Staff, name, trainingStart, trainingEnd, validFrom, validTo
				, reference, outcome, didNotAttend)
			VALUES (@staffID, @name, @trainingStart, @trainingEnd, @validFrom, @validTo
				, @reference, @outcome, @didNotAttend);

		SELECT @ID = MAX(ID_Skill) FROM fusion.staffSkill;

	END
	ELSE
	BEGIN
		UPDATE fusion.staffSkill SET name = @name, validFrom = @validFrom, validTo = @validTo
					, trainingStart = @trainingStart, trainingEnd = @trainingEnd
					, reference = @reference, outcome = @outcome, didNotAttend = @didNotAttend
					WHERE ID_Skill = @ID;
	END

END

GO

CREATE PROCEDURE fusion.pMessageUpdate_StaffPictureChange(@ID int OUTPUT
	, @picture			varbinary(MAX)
)
AS
BEGIN

	DECLARE @photostring varchar(MAX);

	SET @photostring = '<<V002>>2 Embedded Photograph.jpg' + SPACE(367) + convert(varchar(MAX),@picture)

	UPDATE fusion.staff SET picture = convert(varbinary(max), @photostring)
		WHERE StaffID = @ID;

END

GO


IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[fusion].[spSendFusionMessage]') AND xtype = 'P')
	DROP PROCEDURE [fusion].[spSendFusionMessage];
GO

CREATE PROCEDURE fusion.[spSendFusionMessage](@TableID integer, @RecordID integer)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE @messageName varchar(255),
			@parentID	integer;

	-- Personnel Records
	IF @TableID = 1
	BEGIN
		EXEC fusion.[pSendFusionMessageCheckContext] @MessageType='StaffChange', @LocalId=@RecordID;
		EXEC fusion.[pSendFusionMessageCheckContext] @MessageType='StaffPictureChange', @LocalId=@RecordID;
	END

	-- Address 
	IF @TableID = 204
	BEGIN
		SELECT @parentID = ID_1 FROM dbo.Address WHERE id = @RecordID;
		EXEC fusion.[pSendFusionMessageCheckContext] @MessageType='StaffChange', @LocalId=@parentID;
	END

	-- Legal Documents
	IF @TableID = 210
	BEGIN
		SELECT @parentID = ID_1 FROM dbo.Legal_Documents WHERE id = @RecordID;
		EXEC fusion.[pSendFusionMessageCheckContext] @MessageType='StaffChange', @LocalId=@parentID;
		EXEC fusion.[pSendFusionMessageCheckContext] @MessageType='StaffLegalDocumentChange', @LocalId=@RecordID;
	END

	-- Contacts
	IF @TableID = 42
	BEGIN
		SELECT @parentID = ID_1 FROM dbo.Contacts WHERE id = @RecordID;
		EXEC fusion.[pSendFusionMessageCheckContext] @MessageType='StaffChange', @LocalId=@parentID;
		EXEC fusion.[pSendFusionMessageCheckContext] @MessageType='StaffContactChange', @LocalId=@RecordID;
	END

	-- Training Booking
	IF @TableID = 29
	BEGIN
		SELECT @parentID = ID_1 FROM dbo.Training_Booking WHERE id = @RecordID;
		EXEC fusion.[pSendFusionMessageCheckContext] @MessageType='StaffChange', @LocalId=@parentID;
		EXEC fusion.[pSendFusionMessageCheckContext] @MessageType='StaffSkillChange', @LocalId=@RecordID;
	END

	-- Contract
	IF @TableID = 3
	BEGIN
		SELECT @parentID = ID_1 FROM dbo.Appointments WHERE id = @RecordID;
		EXEC fusion.[pSendFusionMessageCheckContext] @MessageType='StaffChange', @LocalId=@parentID;
		EXEC fusion.[pSendFusionMessageCheckContext] @MessageType='StaffContractChange', @LocalId=@RecordID;
	END


END
