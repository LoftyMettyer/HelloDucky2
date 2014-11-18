/* Required Tables

250	- Absence_Entry
251	- Appointment_Absence_Entry
252	- Absence_Breakdown


*/


IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[trcustom_Absence_Entry_P&E]') AND xtype in (N'TR'))
	DROP TRIGGER [dbo].[trcustom_Absence_Entry_P&E]
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[trcustom_Appointment_Absence_Entry_P&E]') AND xtype in (N'TR'))
	DROP TRIGGER [dbo].[trcustom_Appointment_Absence_Entry_P&E]
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[trcustom_Absence_Breakdown_P&E]') AND xtype in (N'TR'))
	DROP TRIGGER [dbo].[trcustom_Absence_Breakdown_P&E]
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[trcustom_Absence_Breakdown_P&E_D02]') AND xtype in (N'TR'))
	DROP TRIGGER [dbo].[trcustom_Absence_Breakdown_P&E_D02]
GO


-- Some system triggers that need disabling/removing
DISABLE TRIGGER trsys_Absence_i01 ON [dbo].[tbuser_Absence]
GO

DISABLE TRIGGER trsys_Absence_i02 ON [dbo].[tbuser_Absence]
GO

DISABLE TRIGGER trsys_Absence_u01 ON [dbo].[tbuser_Absence]
GO

DISABLE TRIGGER trsys_Absence_u02 ON [dbo].[tbuser_Absence]
GO

DISABLE TRIGGER trsys_Absence_d01 ON [dbo].[tbuser_Absence]
GO


DISABLE TRIGGER trsys_Absence_Entry_i01 ON [dbo].[tbuser_Absence_Entry]
GO

DISABLE TRIGGER trsys_Absence_Entry_i02 ON [dbo].[tbuser_Absence_Entry]
GO

DISABLE TRIGGER trsys_Absence_Entry_u01 ON [dbo].[tbuser_Absence_Entry]
GO

DISABLE TRIGGER trsys_Absence_Entry_u02 ON [dbo].[tbuser_Absence_Entry]
GO

DISABLE TRIGGER trsys_Absence_Entry_d01 ON [dbo].[tbuser_Absence_Entry]
GO

DISABLE TRIGGER trsys_Appointment_Absence_Entry_i01 ON [dbo].[tbuser_Appointment_Absence_Entry]
GO

DISABLE TRIGGER trsys_Appointment_Absence_Entry_i02 ON [dbo].[tbuser_Appointment_Absence_Entry]
GO

DISABLE TRIGGER trsys_Appointment_Absence_Entry_u01 ON [dbo].[tbuser_Appointment_Absence_Entry]
GO

DISABLE TRIGGER trsys_Appointment_Absence_Entry_u02 ON [dbo].[tbuser_Appointment_Absence_Entry]
GO

DISABLE TRIGGER trsys_Appointment_Absence_Entry_d01 ON [dbo].[tbuser_Appointment_Absence_Entry]
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[trsys_Absence_Breakdown]') AND xtype in (N'TR'))
	DROP TRIGGER [dbo].[trsys_Absence_Breakdown];
GO


CREATE TRIGGER [dbo].[trcustom_Absence_Entry_P&E] ON [dbo].[tbuser_Absence_Entry]
    AFTER INSERT, UPDATE, DELETE
AS
BEGIN
    SET NOCOUNT ON;

    DELETE [dbo].[tbuser_Absence_Breakdown] WHERE [id_250] IN (SELECT DISTINCT [id] FROM deleted);

	INSERT Absence_Breakdown([source], ID_250, Post_ID, [Type], Payroll_Type_Code, Reason, Payroll_Reason_Code, Absence_In, Duration, Absence_Date, [Session]
		, Day_Pattern_AM, Day_Pattern_PM, Hour_Pattern_AM, Hour_Pattern_PM, Staff_Number, Payroll_Company_Code)	
		SELECT 'pers', i.ID, ap.ID, i.Absence_Type, i.Payroll_Type_Code, i.Reason, i.Payroll_Reason_Code, wp.Absence_In
			, CASE 
				WHEN DATEPART(dw,dr.IndividualDate) = 1 AND dr.SessionType = 'AM' THEN wp.Sunday_Hours_AM
				WHEN DATEPART(dw,dr.IndividualDate) = 2 AND dr.SessionType = 'AM' THEN wp.Monday_Hours_AM
				WHEN DATEPART(dw,dr.IndividualDate) = 3 AND dr.SessionType = 'AM' THEN wp.Tuesday_Hours_AM
				WHEN DATEPART(dw,dr.IndividualDate) = 4 AND dr.SessionType = 'AM' THEN wp.Wednesday_Hours_AM
				WHEN DATEPART(dw,dr.IndividualDate) = 5 AND dr.SessionType = 'AM' THEN wp.Thursday_Hours_AM
				WHEN DATEPART(dw,dr.IndividualDate) = 6 AND dr.SessionType = 'AM' THEN wp.Friday_Hours_AM
				WHEN DATEPART(dw,dr.IndividualDate) = 7 AND dr.SessionType = 'AM' THEN wp.Saturday_Hours_AM
				WHEN DATEPART(dw,dr.IndividualDate) = 1 AND dr.SessionType = 'PM' THEN wp.Sunday_Hours_PM
				WHEN DATEPART(dw,dr.IndividualDate) = 2 AND dr.SessionType = 'PM' THEN wp.Monday_Hours_PM
				WHEN DATEPART(dw,dr.IndividualDate) = 3 AND dr.SessionType = 'PM' THEN wp.Tuesday_Hours_PM
				WHEN DATEPART(dw,dr.IndividualDate) = 4 AND dr.SessionType = 'PM' THEN wp.Wednesday_Hours_PM
				WHEN DATEPART(dw,dr.IndividualDate) = 5 AND dr.SessionType = 'PM' THEN wp.Thursday_Hours_PM
				WHEN DATEPART(dw,dr.IndividualDate) = 6 AND dr.SessionType = 'PM' THEN wp.Friday_Hours_PM
				WHEN DATEPART(dw,dr.IndividualDate) = 7 AND dr.SessionType = 'PM' THEN wp.Saturday_Hours_PM
				WHEN DATEPART(dw,dr.IndividualDate) = 1 AND dr.SessionType = 'Day' THEN wp.Sunday_Hours_AM + Sunday_Hours_PM
				WHEN DATEPART(dw,dr.IndividualDate) = 2 AND dr.SessionType = 'Day' THEN wp.Monday_Hours_AM + Monday_Hours_PM
				WHEN DATEPART(dw,dr.IndividualDate) = 3 AND dr.SessionType = 'Day' THEN wp.Tuesday_Hours_AM + Tuesday_Hours_PM
				WHEN DATEPART(dw,dr.IndividualDate) = 4 AND dr.SessionType = 'Day' THEN wp.Wednesday_Hours_AM + Wednesday_Hours_PM
				WHEN DATEPART(dw,dr.IndividualDate) = 5 AND dr.SessionType = 'Day' THEN wp.Thursday_Hours_AM + Thursday_Hours_PM
				WHEN DATEPART(dw,dr.IndividualDate) = 6 AND dr.SessionType = 'Day' THEN wp.Friday_Hours_AM + Friday_Hours_PM
				WHEN DATEPART(dw,dr.IndividualDate) = 7 AND dr.SessionType = 'Day' THEN wp.Saturday_Hours_AM + Saturday_Hours_PM
			END
			, dr.IndividualDate, dr.SessionType
			, wp.Day_Pattern_AM, wp.Day_Pattern_PM, wp.Hour_Pattern_AM, wp.Hour_Pattern_PM
			, pr.Staff_Number, pr.Payroll_Company_Code
		FROM inserted i
			CROSS APPLY [dbo].[udfsysDateRangeToTable] ('d', i.Start_Date, i.Start_Session,  i.End_Date, i.End_Session) dr
			INNER JOIN Appointments ap ON ap.ID_1 = i.ID_1
			INNER JOIN Appointment_Working_Patterns wp ON wp.ID_3 = ap.ID
			INNER JOIN Personnel_Records pr ON pr.ID = i.ID_1
		WHERE wp.Effective_Date <= dr.IndividualDate AND (wp.End_Date >= dr.IndividualDate OR wp.End_Date IS NULL);

END
GO

CREATE TRIGGER [dbo].[trcustom_Appointment_Absence_Entry_P&E] ON [dbo].[tbuser_Appointment_Absence_Entry]
    AFTER INSERT, UPDATE, DELETE
AS
BEGIN
    SET NOCOUNT ON;

    DELETE [dbo].[tbuser_Absence_Breakdown] WHERE [id_251] IN (SELECT DISTINCT [id] FROM deleted);

	INSERT Absence_Breakdown([source], ID_251, Post_ID, [Type], Payroll_Type_Code, Reason, Payroll_Reason_Code, Absence_In, Duration, Absence_Date, [Session]
		, Day_Pattern_AM, Day_Pattern_PM, Hour_Pattern_AM, Hour_Pattern_PM, Staff_Number, Payroll_Company_Code)	
		SELECT 'post', i.ID, wp.ID_3, i.Absence_Type, i.Payroll_Type_Code, i.Reason, i.Payroll_Reason_Code, wp.Absence_In
			, CASE 
				WHEN DATEPART(dw,dr.IndividualDate) = 1 AND dr.SessionType = 'AM' THEN wp.Sunday_Hours_AM
				WHEN DATEPART(dw,dr.IndividualDate) = 2 AND dr.SessionType = 'AM' THEN wp.Monday_Hours_AM
				WHEN DATEPART(dw,dr.IndividualDate) = 3 AND dr.SessionType = 'AM' THEN wp.Tuesday_Hours_AM
				WHEN DATEPART(dw,dr.IndividualDate) = 4 AND dr.SessionType = 'AM' THEN wp.Wednesday_Hours_AM
				WHEN DATEPART(dw,dr.IndividualDate) = 5 AND dr.SessionType = 'AM' THEN wp.Thursday_Hours_AM
				WHEN DATEPART(dw,dr.IndividualDate) = 6 AND dr.SessionType = 'AM' THEN wp.Friday_Hours_AM
				WHEN DATEPART(dw,dr.IndividualDate) = 7 AND dr.SessionType = 'AM' THEN wp.Saturday_Hours_AM
				WHEN DATEPART(dw,dr.IndividualDate) = 1 AND dr.SessionType = 'PM' THEN wp.Sunday_Hours_PM
				WHEN DATEPART(dw,dr.IndividualDate) = 2 AND dr.SessionType = 'PM' THEN wp.Monday_Hours_PM
				WHEN DATEPART(dw,dr.IndividualDate) = 3 AND dr.SessionType = 'PM' THEN wp.Tuesday_Hours_PM
				WHEN DATEPART(dw,dr.IndividualDate) = 4 AND dr.SessionType = 'PM' THEN wp.Wednesday_Hours_PM
				WHEN DATEPART(dw,dr.IndividualDate) = 5 AND dr.SessionType = 'PM' THEN wp.Thursday_Hours_PM
				WHEN DATEPART(dw,dr.IndividualDate) = 6 AND dr.SessionType = 'PM' THEN wp.Friday_Hours_PM
				WHEN DATEPART(dw,dr.IndividualDate) = 7 AND dr.SessionType = 'PM' THEN wp.Saturday_Hours_PM
				WHEN DATEPART(dw,dr.IndividualDate) = 1 AND dr.SessionType = 'Day' THEN wp.Sunday_Hours_AM + Sunday_Hours_PM
				WHEN DATEPART(dw,dr.IndividualDate) = 2 AND dr.SessionType = 'Day' THEN wp.Monday_Hours_AM + Monday_Hours_PM
				WHEN DATEPART(dw,dr.IndividualDate) = 3 AND dr.SessionType = 'Day' THEN wp.Tuesday_Hours_AM + Tuesday_Hours_PM
				WHEN DATEPART(dw,dr.IndividualDate) = 4 AND dr.SessionType = 'Day' THEN wp.Wednesday_Hours_AM + Wednesday_Hours_PM
				WHEN DATEPART(dw,dr.IndividualDate) = 5 AND dr.SessionType = 'Day' THEN wp.Thursday_Hours_AM + Thursday_Hours_PM
				WHEN DATEPART(dw,dr.IndividualDate) = 6 AND dr.SessionType = 'Day' THEN wp.Friday_Hours_AM + Friday_Hours_PM
				WHEN DATEPART(dw,dr.IndividualDate) = 7 AND dr.SessionType = 'Day' THEN wp.Saturday_Hours_AM + Saturday_Hours_PM
			END
			, dr.IndividualDate, dr.SessionType
			, wp.Day_Pattern_AM, wp.Day_Pattern_PM, wp.Hour_Pattern_AM, wp.Hour_Pattern_PM
			, pr.Staff_Number, pr.Payroll_Company_Code
		FROM inserted i
			CROSS APPLY [dbo].[udfsysDateRangeToTable] ('d', i.Start_Date, i.Start_Session,  i.End_Date, i.End_Session) dr
			INNER JOIN Appointments ap ON ap.ID = i.ID_3
			INNER JOIN Appointment_Working_Patterns wp ON wp.ID_3 = i.ID_3
			INNER JOIN Personnel_Records pr ON pr.ID = ap.ID_1
		WHERE wp.Effective_Date <= dr.IndividualDate AND (wp.End_Date >= dr.IndividualDate OR wp.End_Date IS NULL);


END
GO


CREATE TRIGGER [dbo].[trcustom_Absence_Breakdown_P&E] ON [dbo].[tbuser_Absence_Breakdown]
    AFTER INSERT
AS
BEGIN
    SET NOCOUNT ON;

	DECLARE @AbsenceID	integer,
			@startDate	datetime,
			@endDate	datetime;

	INSERT Absence(Absence_Type, Payroll_Code, Reason, Payroll_Reason, Start_Date, Start_Session, End_Date, End_Session, Absence_In, Duration_Days, Duration_Hours)
		SELECT DISTINCT ab.Type, ab.Payroll_Type_Code, ab.Reason, ab.Payroll_Reason_Code
			, m.startdate
			, (SELECT DISTINCT CASE WHEN [Session] = 'Day' THEN 'AM' ELSE [Session] END FROM inserted WHERE (ID_250 = ab.ID_250 OR ID_251 = ab.ID_251) AND Absence_Date = m.startdate)
			, m.enddate
			, (SELECT DISTINCT CASE WHEN [Session] = 'Day' THEN 'PM' ELSE [Session] END FROM inserted WHERE (ID_250 = ab.ID_250 OR ID_251 = ab.ID_251) AND Absence_Date = m.enddate)
			, '???' , 0, m.Duration
		FROM inserted ab
		CROSS APPLY (
			SELECT MIN(range.Absence_Date) AS startdate, MAX(range.Absence_Date) AS enddate, SUM(Duration) AS Duration
			FROM inserted range
			WHERE ab.ID_250 = range.ID_250 OR ab.ID_251 = range.ID_251) m;


	SELECT @AbsenceID = MAX(ID) FROM Absence

	UPDATE [tbuser_Absence_Breakdown]
		SET id_2 = @AbsenceID
	FROM [inserted] base WHERE base.[id] = [dbo].[tbuser_Absence_Breakdown].[id]

END

GO

CREATE TRIGGER [dbo].[trcustom_Absence_Breakdown_P&E_D02] ON [dbo].[tbuser_Absence_Breakdown]
    INSTEAD OF DELETE
AS
BEGIN
    SET NOCOUNT ON;

	DECLARE @AbsenceID	integer,
			@startDate	datetime,
			@endDate	datetime;

	--SELECT @startDate = MIN(absence_date), @endDate = MAX(absence_date) FROM inserted;

	--SELECT DISTINCT [id_2] FROM deleted;
	--select * from deleted;

    DELETE [dbo].[tbuser_Absence] WHERE [id] IN (SELECT DISTINCT [id_2] FROM deleted);

	WITH base AS (SELECT * FROM dbo.[tbuser_Absence_Breakdown]
        WHERE [id] IN (SELECT DISTINCT [id] FROM deleted))
        DELETE FROM base;



	--INSERT Absence(Start_Date, End_Date)
	--	VALUES (@startDate, @endDate);

	--SELECT @AbsenceID = MAX(ID) FROM Absence

	--UPDATE [tbuser_Absence_Breakdown] SET id_2 = @AbsenceID;

END








GO


--EXEC sp_settriggerorder @triggername=N'[dbo].[trcustom_Absence_Entry_P&E]', @order=N'Last', @stmttype=N'INSERT'
--EXEC sp_settriggerorder @triggername=N'[dbo].[trcustom_Absence_Entry_P&E]', @order=N'Last', @stmttype=N'UPDATE'
--EXEC sp_settriggerorder @triggername=N'[dbo].[trcustom_Absence_Entry_P&E]', @order=N'Last', @stmttype=N'DELETE'

GO


