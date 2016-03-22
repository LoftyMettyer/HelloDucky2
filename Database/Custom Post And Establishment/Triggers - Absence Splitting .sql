
/* Required Tables

1   - Personnel_Records
3   - Appointments
228 - Establishment_Posts
242 - Appointment Allowances
243 - Appointment Benefits
244 - Appointment Deductions
249 - Appointment_Working_Patterns
250	- Absence_Entry
251	- Appointment_Absence
252	- Absence_Breakdown
254 - Appointment_Absence_Staging

*/

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfcustom_AbsenceDurationForAppointment]') AND xtype = 'FN')
	DROP FUNCTION [dbo].[udfcustom_AbsenceDurationForAppointment];

EXEC sp_executesql N'CREATE FUNCTION udfcustom_AbsenceDurationForAppointment(@startDate datetime, @startSession varchar(2), @endDate datetime, @endSession varchar(2), @appointmentID integer)
RETURNS numeric(10,2)
AS
BEGIN

	-- Get change dates for this appointment
	DECLARE @changeDates TABLE(type varchar(1), Effective_Date datetime, AppointmentID integer);
	DECLARE @working_patterns TABLE (type varchar(1),Effective_Date datetime, End_Date datetime
							, Sunday_Hours_AM numeric(4,2), Sunday_Hours_PM numeric(4,2)
							, Monday_Hours_AM numeric(4,2), Monday_Hours_PM numeric(4,2)
							, Tuesday_Hours_AM numeric(4,2), Tuesday_Hours_PM numeric(4,2)
							, Wednesday_Hours_AM numeric(4,2), Wednesday_Hours_PM numeric(4,2)
							, Thursday_Hours_AM numeric(4,2), Thursday_Hours_PM numeric(4,2)
							, Friday_Hours_AM numeric(4,2), Friday_Hours_PM numeric(4,2)
							, Saturday_Hours_AM numeric(4,2), Saturday_Hours_PM numeric(4,2)
							, Day_Pattern_AM varchar(28), Day_Pattern_PM varchar(28)
							, Hour_Pattern_AM varchar(28), Hour_Pattern_PM varchar(28));
	DECLARE @absenceIn varchar(5),
			@duration numeric(10,2);

	SELECT @absenceIn = Absence_In FROM Appointments WHERE ID = @appointmentID;

	SET @startDate = DATEADD(dd, 0, DATEDIFF(dd, 0, @startDate));
	SET @endDate = DATEADD(dd, 0, DATEDIFF(dd, 0, @endDate));

	INSERT @changeDates 
		SELECT ''a'', awp.Effective_Date, a.ID FROM Appointments a
			INNER JOIN Appointment_Working_Patterns awp ON awp.ID_3 = a.ID
		WHERE a.ID = @appointmentID 
	  UNION
		SELECT ''b'', awp.End_Date + 1, a.ID FROM Appointments a
			INNER JOIN Appointment_Working_Patterns awp ON awp.ID_3 = a.ID
		WHERE a.ID = @appointmentID 
			AND awp.End_Date IS NOT NULL AND (a.Appointment_End_Date >= @startDate OR a.Appointment_End_Date IS NULL);

	INSERT @working_patterns
		SELECT cd.type, cd.Effective_Date, NULL
			, ISNULL(SUM(wp.Sunday_Hours_AM),0), ISNULL(SUM(wp.Sunday_Hours_PM),0)
			, ISNULL(SUM(wp.Monday_Hours_AM),0), ISNULL(SUM(wp.Monday_Hours_PM),0)
			, ISNULL(SUM(wp.Tuesday_Hours_AM),0), ISNULL(SUM(wp.Tuesday_Hours_PM),0)
			, ISNULL(SUM(wp.Wednesday_Hours_AM),0), ISNULL(SUM(wp.Wednesday_Hours_PM),0)
			, ISNULL(SUM(wp.Thursday_Hours_AM),0), ISNULL(SUM(wp.Thursday_Hours_PM),0)
			, ISNULL(SUM(wp.Friday_Hours_AM),0), ISNULL(SUM(wp.Friday_Hours_PM),0)
			, ISNULL(SUM(wp.Saturday_Hours_AM),0), ISNULL(SUM(wp.Saturday_Hours_PM),0)
			, NULL, NULL, NULL, NULL
		FROM @changeDates cd
		LEFT JOIN Appointment_Working_Patterns wp ON cd.AppointmentID = wp.id_3 AND cd.Effective_Date >= wp.Effective_Date AND (cd.Effective_Date <= wp.End_Date OR wp.End_Date IS NULL)
			GROUP BY cd.Effective_Date, cd.type;

	UPDATE t 
		SET End_Date = (SELECT top 1 m.Effective_Date - 1 FROM @working_patterns m WHERE m.Effective_Date > t.Effective_Date ORDER BY m.Effective_Date),
			Day_Pattern_AM = dbo.udfsysPatternFromHours (''Days'', Sunday_Hours_AM, Monday_Hours_AM, Tuesday_Hours_AM, Wednesday_Hours_AM, Thursday_Hours_AM, Friday_Hours_AM, Saturday_Hours_AM),
			Day_Pattern_PM = dbo.udfsysPatternFromHours (''Days'', Sunday_Hours_PM, Monday_Hours_PM, Tuesday_Hours_PM, Wednesday_Hours_PM, Thursday_Hours_PM, Friday_Hours_PM, Saturday_Hours_PM),
			Hour_Pattern_AM = dbo.udfsysPatternFromHours (''Hours'', Sunday_Hours_AM, Monday_Hours_AM, Tuesday_Hours_AM, Wednesday_Hours_AM, Thursday_Hours_AM, Friday_Hours_AM, Saturday_Hours_AM),
			Hour_Pattern_PM = dbo.udfsysPatternFromHours (''Hours'', Sunday_Hours_PM, Monday_Hours_PM, Tuesday_Hours_PM, Wednesday_Hours_PM, Thursday_Hours_PM, Friday_Hours_PM, Saturday_Hours_PM)
	FROM @working_patterns t;

	SELECT @duration = SUM(dbo.udfsysDurationFromPattern(@absenceIn, dr.IndividualDate, dr.SessionType, wp.Sunday_Hours_AM, wp.Monday_Hours_AM, wp.Tuesday_Hours_AM
			, wp.Wednesday_Hours_AM, wp.Thursday_Hours_AM, wp.Friday_Hours_AM, wp.Saturday_Hours_AM, wp.Sunday_Hours_PM, wp.Monday_Hours_PM, wp.Tuesday_Hours_PM
			, wp.Wednesday_Hours_PM, wp.Thursday_Hours_PM, wp.Friday_Hours_PM, wp.Saturday_Hours_PM))
	FROM [dbo].[udfsysDateRangeToTable] (''d'', @startDate, @startSession,  @endDate, @endSession) dr
		LEFT JOIN @working_patterns wp ON dr.IndividualDate >= ISNULL(wp.Effective_Date, ''1899-12-31'') AND dr.IndividualDate <= ISNULL(wp.End_Date, ''9999-12-31'')
	WHERE dr.IndividualDate >= @startDate AND dr.IndividualDate <= @endDate;

	RETURN ISNULL(@duration, 0)

END';


DELETE FROM ASRSysTableTriggers
UPDATE ASRSysTables SET deletetriggerdisabled = 0, inserttriggerdisabled = 0 WHERE tableid = 252;


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

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[trsys_Absence_Breakdown]') AND xtype in (N'TR'))
	DROP TRIGGER [dbo].[trsys_Absence_Breakdown];
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[trcustom_Appointments_P&E]') AND xtype in (N'TR'))
	DROP TRIGGER [dbo].[trcustom_Appointments_P&E]
GO

INSERT ASRSysTableTriggers (TriggerID, TableID, Name, CodePosition, IsSystem, Content) VALUES (2, 250, 'Split Absence Request for individual appointment approval', 0, 1, '    
	DECLARE @today datetime =  DATEADD(dd, 0, DATEDIFF(dd, 0, GETDATE()));

	INSERT Appointment_Absence_Staging (ID_3, Start_Date, Start_Session, End_Date, End_Session, Absence_Type, Reason, Duration, Absence_In)
		SELECT a.ID, ae.Start_Date, ae.Start_Session, ae.End_Date, ae.End_Session
				, ae.Absence_Type, ae.Reason
				, dbo.udfcustom_AbsenceDurationForAppointment(ae.Start_Date, ae.Start_Session, ae.End_Date, ae.End_Session, a.ID)
				, a.Absence_In
		FROM inserted ae
		INNER JOIN Appointments a ON ae.ID_1 = a.ID_1 AND (a.post_ID = ae.post_ID OR ae.post_ID IS NULL OR ae.post_ID = '''')
		WHERE ae.Start_Date <= ISNULL(Appointment_End_Date, convert(datetime, ''9999-12-31''))
			AND (ae.End_Date >= ISNULL(Appointment_Start_Date, convert(datetime, ''1899-12-31'')) OR ae.End_Date IS NULL)
			AND dbo.udfcustom_AbsenceDurationForAppointment(ae.Start_Date, ae.Start_Session, ISNULL(ae.End_Date, @today), ae.End_Session, a.ID) > 0;')
GO

INSERT ASRSysTableTriggers (TriggerID, TableID, Name, CodePosition, IsSystem, Content) VALUES (8, 254, 'Manual & Automatic Approval of absence based on type', 1, 1, '    

	-- Recalculate the duration for this absence request
	;WITH base AS (SELECT *
            FROM [dbo].[tbuser_Appointment_Absence_Staging]
            WHERE [id] IN (SELECT DISTINCT [id] FROM inserted))
    UPDATE base SET 
		Duration = dbo.udfcustom_AbsenceDurationForAppointment(Start_Date, Start_Session, End_Date, End_Session, ID_3);

	-- Approve authorised absences and absences that require no authorisation
	INSERT Appointment_Absence (ID_3, StagingID, Start_Date, End_Date, Start_Session, End_Session, Absence_In, Absence_Type, Reason, Post_ID, Staff_Number, Payroll_Company_Code)
		SELECT i.ID_3, i.ID, i.Start_Date, i.End_Date, i.Start_Session, i.End_Session, a.Absence_In, i.Absence_Type, i.Reason, a.Post_ID, a.Staff_Number, a.Payroll_Company_Code
		FROM inserted i
			INNER JOIN Appointments a ON a.ID = i.ID_3
			INNER JOIN Absence_Type_Table at ON at.Absence_Type = i.Absence_Type
		WHERE (at.Requires_Authorisation = 0 OR i.Status = ''Authorised'')
			AND i.ID NOT IN (SELECT StagingID FROM Appointment_Absence);

')
GO

INSERT ASRSysTableTriggers (TriggerID, TableID, Name, CodePosition, IsSystem, Content) VALUES (3, 251, 'Breakdown absences into individual days', 1, 1, '   
   	DECLARE @today datetime =  DATEADD(dd, 0, DATEDIFF(dd, 0, GETDATE()));
	DECLARE @changeDates TABLE(Effective_Date datetime, AppointmentID integer, PersID integer);

	IF NOT EXISTS(SELECT [tablefromid] FROM #intransactiontrigger WHERE [tablefromid] = 252)
	BEGIN
		
		DELETE [dbo].[tbuser_Absence_Breakdown] WHERE [id_251] IN (SELECT DISTINCT [id] FROM deleted);

		INSERT @changeDates 
			SELECT awp.Effective_Date, a.ID, a.ID_1 FROM Appointments a
				INNER JOIN inserted e ON e.ID_3 = a.ID
				INNER JOIN Appointment_Working_Patterns awp ON awp.ID_3 = a.ID
			WHERE awp.Effective_Date > a.Appointment_Start_Date
			UNION
			SELECT awp.End_Date + 1, a.ID, a.ID_1 FROM Appointments a
				INNER JOIN inserted e ON e.ID_3 = a.ID
				INNER JOIN Appointment_Working_Patterns awp ON awp.ID_3 = a.ID
			WHERE awp.End_Date IS NOT NULL
			UNION
			SELECT a.Appointment_Start_Date, a.ID, a.ID_1 FROM Appointments a
				INNER JOIN inserted e ON e.ID_3 = a.ID
			UNION
			SELECT a.Appointment_End_Date + 1, a.ID, a.ID_1 FROM Appointments a
				INNER JOIN inserted e ON e.ID_3 = a.ID
			UNION 
			SELECT ''1899-12-31'', a.ID, a.ID_1 FROM Appointments a
				INNER JOIN inserted e ON e.ID_3 = a.ID;

		DECLARE @working_patterns TABLE (PersID integer, Effective_Date datetime, End_Date datetime, AppointmentID integer
								, Sunday_Hours_AM numeric(4,2), Sunday_Hours_PM numeric(4,2)
								, Monday_Hours_AM numeric(4,2), Monday_Hours_PM numeric(4,2)
								, Tuesday_Hours_AM numeric(4,2), Tuesday_Hours_PM numeric(4,2)
								, Wednesday_Hours_AM numeric(4,2), Wednesday_Hours_PM numeric(4,2)
								, Thursday_Hours_AM numeric(4,2), Thursday_Hours_PM numeric(4,2)
								, Friday_Hours_AM numeric(4,2), Friday_Hours_PM numeric(4,2)
								, Saturday_Hours_AM numeric(4,2), Saturday_Hours_PM numeric(4,2)
								, Day_Pattern_AM varchar(28), Day_Pattern_PM varchar(28)
								, Hour_Pattern_AM varchar(28), Hour_Pattern_PM varchar(28));

		INSERT @working_patterns
			SELECT DISTINCT cd.PersID, cd.Effective_Date, NULL, cd.AppointmentID
				, ISNULL(wp.Sunday_Hours_AM,0), ISNULL(wp.Sunday_Hours_PM,0)
				, ISNULL(wp.Monday_Hours_AM,0), ISNULL(wp.Monday_Hours_PM,0)
				, ISNULL(wp.Tuesday_Hours_AM,0), ISNULL(wp.Tuesday_Hours_PM,0)
				, ISNULL(wp.Wednesday_Hours_AM,0), ISNULL(wp.Wednesday_Hours_PM,0)
				, ISNULL(wp.Thursday_Hours_AM,0), ISNULL(wp.Thursday_Hours_PM,0)
				, ISNULL(wp.Friday_Hours_AM,0), ISNULL(wp.Friday_Hours_PM,0)
				, ISNULL(wp.Saturday_Hours_AM,0), ISNULL(wp.Saturday_Hours_PM,0)
				, NULL, NULL, NULL, NULL
			FROM @changeDates cd
			LEFT JOIN Appointment_Working_Patterns wp ON cd.AppointmentID = wp.id_3 AND cd.Effective_Date = wp.Effective_Date

		UPDATE t 
			SET End_Date = (SELECT top 1 m.Effective_Date - 1 FROM @working_patterns m WHERE m.Effective_Date > t.Effective_Date AND m.AppointmentID = t.AppointmentID ORDER BY m.Effective_Date),
				Day_Pattern_AM = dbo.udfsysPatternFromHours (''Days'', Sunday_Hours_AM, Monday_Hours_AM, Tuesday_Hours_AM, Wednesday_Hours_AM, Thursday_Hours_AM, Friday_Hours_AM, Saturday_Hours_AM),
				Day_Pattern_PM = dbo.udfsysPatternFromHours (''Days'', Sunday_Hours_PM, Monday_Hours_PM, Tuesday_Hours_PM, Wednesday_Hours_PM, Thursday_Hours_PM, Friday_Hours_PM, Saturday_Hours_PM),
				Hour_Pattern_AM = dbo.udfsysPatternFromHours (''Hours'', Sunday_Hours_AM, Monday_Hours_AM, Tuesday_Hours_AM, Wednesday_Hours_AM, Thursday_Hours_AM, Friday_Hours_AM, Saturday_Hours_AM),
				Hour_Pattern_PM = dbo.udfsysPatternFromHours (''Hours'', Sunday_Hours_PM, Monday_Hours_PM, Tuesday_Hours_PM, Wednesday_Hours_PM, Thursday_Hours_PM, Friday_Hours_PM, Saturday_Hours_PM)
		FROM @working_patterns t

		DELETE [dbo].[tbuser_Absence_Breakdown] WHERE [id_251] IN (SELECT DISTINCT [id] FROM deleted);

		INSERT Absence_Breakdown(ID_251, Duration, Absence_Date, [Session]
			, Day_Pattern_AM, Day_Pattern_PM, Hour_Pattern_AM, Hour_Pattern_PM)	
			SELECT i.ID
				, dbo.udfsysDurationFromPattern(ap.Absence_In, dr.IndividualDate, dr.SessionType, wp.Sunday_Hours_AM, wp.Monday_Hours_AM, wp.Tuesday_Hours_AM, wp.Wednesday_Hours_AM, wp.Thursday_Hours_AM, wp.Friday_Hours_AM, wp.Saturday_Hours_AM, wp.Sunday_Hours_PM, wp.Monday_Hours_PM, wp.Tuesday_Hours_PM, wp.Wednesday_Hours_PM, wp.Thursday_Hours_PM, wp.Friday_Hours_PM, wp.Saturday_Hours_PM)
				, dr.IndividualDate, dr.SessionType
				, wp.Day_Pattern_AM, wp.Day_Pattern_PM, wp.Hour_Pattern_AM, wp.Hour_Pattern_PM
			FROM inserted i
				CROSS APPLY [dbo].[udfsysDateRangeToTable] (''d'', i.Start_Date, i.Start_Session, i.End_Date, i.End_Session) dr
				INNER JOIN Appointments ap ON ap.ID = i.ID_3
				INNER JOIN @working_patterns wp ON wp.AppointmentID = ap.ID
				INNER JOIN Personnel_Records pr ON pr.ID = ap.ID_1
				LEFT JOIN Absence_Type_Table at ON at.Absence_Type = i.Absence_Type
				LEFT JOIN Absence_Reason_Table ar ON ar.Reason = i.Reason
			WHERE wp.Effective_Date <= dr.IndividualDate AND (wp.End_Date >= dr.IndividualDate OR wp.End_Date IS NULL)
				AND dbo.udfsysDurationFromPattern(ap.Absence_In, dr.IndividualDate, dr.SessionType, wp.Sunday_Hours_AM, wp.Monday_Hours_AM, wp.Tuesday_Hours_AM, wp.Wednesday_Hours_AM, wp.Thursday_Hours_AM, wp.Friday_Hours_AM, wp.Saturday_Hours_AM, wp.Sunday_Hours_PM, wp.Monday_Hours_PM, wp.Tuesday_Hours_PM, wp.Wednesday_Hours_PM, wp.Thursday_Hours_PM, wp.Friday_Hours_PM, wp.Saturday_Hours_PM) > 0;

	END
			
	-- Update the duration
	MERGE Appointment_Absence TARGET
	USING (SELECT id_251, SUM(Duration) AS Duration FROM Absence_Breakdown
		WHERE Id_251 in (SELECT id FROM inserted)
		GROUP BY ID_251)
	AS SOURCE ON SOURCE.ID_251 = TARGET.ID
	WHEN MATCHED
		THEN UPDATE SET Duration = SOURCE.Duration;

	-- Merge into the personnel absence table
	MERGE Absence AS TARGET
		USING (SELECT aa.Absence_Type, aa.Reason, aa.Absence_In, aa.start_date, aa.Start_Session, aa.end_date, aa.End_Session, aa.Duration
				, a.Post_ID, a.id_1, aa.ID
			FROM Appointment_Absence aa
			INNER JOIN inserted i ON i.ID = aa.ID
			INNER JOIN Appointments a ON a.ID = aa.ID_3)
	AS ins ON ins.ID = TARGET.Appointment_Absence_ID
	WHEN NOT MATCHED
		THEN INSERT (Absence_Type, Reason, Absence_In, Start_Date, Start_Session, End_Date, End_Session, Duration, Post_ID, ID_1, Appointment_Absence_ID)		
			VALUES(ins.Absence_Type, ins.Reason, ins.Absence_In, ins.Start_Date, ins.Start_Session, ins.End_Date, ins.End_Session, Duration, ins.Post_ID, ins.ID_1, ins.ID)
	WHEN MATCHED 
		THEN UPDATE SET Absence_Type = ins.Absence_Type, Reason = ins.Reason, Absence_in = ins.Absence_in, Start_Date = ins.Start_Date, Start_session = ins.Start_Session
			, End_Date = ins.End_Date, Duration = ins.Duration, Post_ID = ins.Post_ID, ID_1 = ins.ID_1
	WHEN NOT MATCHED BY SOURCE AND TARGET.Appointment_Absence_ID IN (SELECT id FROM inserted)
		THEN DELETE;
			
	');
		


INSERT ASRSysTableTriggers (TriggerID, TableID, Name, CodePosition, IsSystem, Content) VALUES (5, 3, 'Populate from Post Template', 0, 1, '    INSERT Appointment_Allowances(ID_3, Effective_Date, Type, Frequency, Amount, Currency)
		SELECT i.ID, i.Appointment_Start_Date, pa.Type, pa.Frequency, pa.Amount, pa.Currency
			FROM inserted i
			INNER JOIN Post_Allowances pa ON pa.ID_219 = i.ID_219;

	INSERT Appointment_Benefits(ID_3, Effective_Date, Type, Provider, Frequency, Cost, Currency, Annual_Cost_to_Employer)
		SELECT i.ID, i.Appointment_Start_Date, pb.Type, pb.Provider, pb.Frequency, pb.Cost, pb.Currency, pb.Annual_Cost_to_Employer
			FROM inserted i
			INNER JOIN Post_Benefits pb ON pb.ID_219 = i.ID_219;

	INSERT Appointment_Deductions(ID_3, Effective_Date, Type, Frequency, Amount, Currency)
		SELECT i.ID, i.Appointment_Start_Date, pd.Type, pd.Frequency, pd.Amount, pd.Currency
			FROM inserted i
			INNER JOIN Post_Deductions pd ON pd.ID_219 = i.ID_219;

	INSERT Appointment_Leave_Schemes(ID_3, Effective_Date, Leave_Scheme)
		SELECT i.ID, i.Appointment_Start_Date, phs.Leave_Scheme
			FROM inserted i
			INNER JOIN Post_Leave_Schemes phs ON phs.ID_219 = i.ID_219;

	INSERT Appointment_OMP_Schemes(ID_3, Effective_Date, OMP_Scheme)
		SELECT i.ID, i.Appointment_Start_Date, sch.OMP_Scheme
			FROM inserted i
			INNER JOIN Post_OMP_Schemes sch ON sch.ID_219 = i.ID_219;

	INSERT Appointment_OSP_Schemes(ID_3, Effective_Date, OSP_Scheme)
		SELECT i.ID, i.Appointment_Start_Date, sch.OSP_Scheme
			FROM inserted i
			INNER JOIN Post_OSP_Schemes sch ON sch.ID_219 = i.ID_219;

	INSERT Appointment_Pension_Schemes(ID_3, Effective_Date, Pension_Scheme)
		SELECT i.ID, i.Appointment_Start_Date, sch.Pension_Scheme
			FROM inserted i
			INNER JOIN Post_Pension_Schemes sch ON sch.ID_219 = i.ID_219;

	INSERT Appointment_Working_Patterns(ID_3, Effective_Date, Regional_ID, Absence_In, Day_Pattern
									, Sunday_Hours_AM, Sunday_Hours_PM, Monday_Hours_AM, Monday_Hours_PM
									, Tuesday_Hours_AM, Tuesday_Hours_PM, Wednesday_Hours_AM, Wednesday_Hours_PM, Thursday_Hours_AM, Thursday_Hours_PM
									, Friday_Hours_AM, Friday_Hours_PM, Saturday_Hours_AM, Saturday_Hours_PM)
		SELECT i.ID, i.Appointment_Start_Date, wp.Regional_ID, wp.Absence_In, wp.Day_Pattern, wp.Sunday_Hours_AM, wp.Sunday_Hours_PM, wp.Monday_Hours_AM,wp.Monday_Hours_PM
									, wp.Tuesday_Hours_AM, wp.Tuesday_Hours_PM, wp.Wednesday_Hours_AM, wp.Wednesday_Hours_PM, wp.Thursday_Hours_AM, wp.Thursday_Hours_PM
									, wp.Friday_Hours_AM, wp.Friday_Hours_PM, wp.Saturday_Hours_AM, wp.Saturday_Hours_PM
			FROM inserted i
			INNER JOIN Post_Working_Patterns wp ON wp.ID_219 = i.ID_219;');
GO

INSERT ASRSysTableTriggers (TriggerID, TableID, Name, CodePosition, IsSystem, Content) VALUES (6, 249, 'Slave to appointment working pattern', 1, 1, '    

	DECLARE @persID integer;
	DECLARE @changeDates TABLE(effectiveDate datetime, ID integer, PersID integer);
	DECLARE @employees TABLE(PersID integer);

	INSERT @employees 
	  SELECT DISTINCT [ID_1] FROM inserted i	
		INNER JOIN Appointments a ON a.ID = i.ID_3;

	INSERT @changeDates 
	  SELECT awp.Effective_Date, awp.ID , a.ID_1
		FROM Appointment_Working_Patterns awp
		INNER JOIN Appointments a ON a.ID = awp.ID_3	
		INNER JOIN @employees e ON e.PersID = a.ID_1
	  UNION
	  SELECT awp.End_Date + 1, awp.ID, a.ID_1
		FROM Appointment_Working_Patterns awp
		INNER JOIN Appointments a ON a.ID = awp.ID_3
		INNER JOIN @employees e ON e.PersID = a.ID_1
		WHERE awp.End_Date IS NOT NULL;


	DECLARE @merged TABLE (PersID integer, effective_date datetime, sunHoursAM numeric(4,2), sunHoursPM numeric(4,2)
						, MonHoursAM numeric(4,2), MonHoursPM numeric(4,2)
						, TuesHoursAM numeric(4,2), TuesHoursPM numeric(4,2)
						, WedHoursAM numeric(4,2), WedHoursPM numeric(4,2)
						, ThursHoursAM numeric(4,2), ThursHoursPM numeric(4,2)
						, FriHoursAM numeric(4,2), FriHoursPM numeric(4,2)
						, SatHoursAM numeric(4,2), SatHoursPM numeric(4,2));

	DECLARE @cursRollupWorkingPatterns cursor,
		@effectiveDate datetime;

    SET @cursRollupWorkingPatterns = CURSOR LOCAL FAST_FORWARD READ_ONLY FOR 
		SELECT DISTINCT EffectiveDate, PersID FROM @changeDates
			ORDER BY EffectiveDate;

	OPEN @cursRollupWorkingPatterns
	FETCH NEXT FROM @cursRollupWorkingPatterns INTO @effectiveDate, @PersID
    WHILE (@@fetch_status = 0)
	BEGIN

		INSERT @merged
			SELECT @PersID, @effectiveDate
			, ISNULL(SUM(wp.Sunday_Hours_AM),0), ISNULL(SUM(wp.Sunday_Hours_PM),0)
			, ISNULL(SUM(wp.Monday_Hours_AM),0), ISNULL(SUM(wp.Monday_Hours_PM),0)
			, ISNULL(SUM(wp.Tuesday_Hours_AM),0), ISNULL(SUM(wp.Tuesday_Hours_PM),0)
			, ISNULL(SUM(wp.Wednesday_Hours_AM),0), ISNULL(SUM(wp.Wednesday_Hours_PM),0)
			, ISNULL(SUM(wp.Thursday_Hours_AM),0), ISNULL(SUM(wp.Thursday_Hours_PM),0)
			, ISNULL(SUM(wp.Friday_Hours_AM),0), ISNULL(SUM(wp.Friday_Hours_PM),0)
			, ISNULL(SUM(wp.Saturday_Hours_AM),0), ISNULL(SUM(wp.Saturday_Hours_PM),0)
		FROM Appointment_Working_Patterns wp
			INNER JOIN Appointments a ON a.ID = wp.ID_3
			WHERE (a.ID_1 = @persID	AND @effectiveDate >= wp.Effective_Date AND (@effectiveDate <= wp.End_Date OR wp.End_Date IS NULL));

		FETCH NEXT FROM @cursRollupWorkingPatterns INTO @effectiveDate, @PersID
	END
	CLOSE @cursRollupWorkingPatterns;
    DEALLOCATE @cursRollupWorkingPatterns;

	MERGE Working_Patterns AS wp
		USING (SELECT PersID, effective_date, SunHoursAM, SunHoursPM, MonHoursAM, MonHoursPM, TuesHoursAM, TuesHoursPM
			, WedHoursAM, WedHoursPM, ThursHoursAM, ThursHoursPM, FriHoursAM, FriHoursPM, SatHoursAM, SatHoursPM 
		FROM @merged)
	AS awp ON (wp.ID_1 = awp.PersID AND wp.Effective_Date = awp.effective_date) 
	WHEN NOT MATCHED BY TARGET
		THEN INSERT(ID_1, Effective_Date, Sunday_Hours_AM, Sunday_Hours_PM, Monday_Hours_AM, Monday_Hours_PM, Tuesday_Hours_AM, Tuesday_Hours_PM, Wednesday_Hours_AM, Wednesday_Hours_PM, Thursday_Hours_AM, Thursday_Hours_PM, Friday_Hours_AM, Friday_Hours_PM, Saturday_Hours_AM, Saturday_Hours_PM) 
			VALUES(awp.persID, awp.Effective_Date, awp.SunHoursAM, awp.SunHoursPM, awp.MonHoursAM, awp.MonHoursPM, awp.TuesHoursAM, awp.TuesHoursPM, awp.WedHoursAM, awp.WedHoursPM, awp.ThursHoursAM, awp.ThursHoursPM, awp.FriHoursAM, awp.FriHoursPM, awp.SatHoursAM, awp.SatHoursPM)
	WHEN MATCHED 
		THEN UPDATE SET Effective_Date = awp.Effective_Date, Sunday_Hours_AM = awp.SunHoursAM, Sunday_Hours_PM = awp.SunHoursPM, Monday_Hours_AM = awp.MonHoursAM, Monday_Hours_PM = awp.MonHoursPM, Tuesday_Hours_AM = awp.TuesHoursAM, Tuesday_Hours_PM = awp.TuesHoursPM, Wednesday_Hours_AM = awp.WedHoursAM, Wednesday_Hours_PM = awp.WedHoursPM
			, Thursday_Hours_AM = awp.ThursHoursAM, Thursday_Hours_PM = awp.ThursHoursPM, Friday_Hours_AM = awp.FriHoursAM, Friday_Hours_PM = awp.FriHoursPM, Saturday_Hours_AM = awp.SatHoursAM, Saturday_Hours_PM = awp.SatHoursPM
	WHEN NOT MATCHED BY SOURCE AND wp.ID_1 IN(SELECT PersID FROM @employees)
	    THEN DELETE ;');

GO

INSERT ASRSysTableTriggers (TriggerID, TableID, Name, CodePosition, IsSystem, Content) VALUES (7, 219, 'Transfer working pattern from appointment', 0, 1, '    INSERT Post_Leave_Schemes (ID_219, Effective_Date, Leave_Scheme)
		SELECT i.ID, i.Effective_Date, chs.Leave_Scheme
			FROM inserted i
			INNER JOIN Contract_Templates ct ON ct.Contract = i.Contract
			INNER JOIN Contract_Leave_Schemes chs ON chs.ID_215 = ct.ID AND (chs.End_Date >= GETDATE() OR End_Date IS NULL);

	INSERT Post_OMP_Schemes (ID_219, Effective_Date, OMP_Scheme, Description)
		SELECT i.ID, i.Effective_Date, omp.OMP_Scheme, omp.Description
			FROM inserted i
			INNER JOIN Contract_Templates ct ON ct.Contract = i.Contract
			INNER JOIN Contract_OMP_Schemes omp ON omp.ID_215 = ct.ID AND (omp.End_Date >= GETDATE() OR End_Date IS NULL);

	INSERT Post_OSP_Schemes (ID_219, Effective_Date, OSP_Scheme, Description)
		SELECT i.ID, i.Effective_Date, osp.OSP_Scheme, osp.Description
			FROM inserted i
			INNER JOIN Contract_Templates ct ON ct.Contract = i.Contract
			INNER JOIN Contract_OSP_Schemes osp ON osp.ID_215 = ct.ID AND (osp.End_Date >= GETDATE() OR End_Date IS NULL);

	INSERT Post_Pension_Schemes (ID_219, Effective_Date, Pension_Scheme)
		SELECT i.ID, i.Effective_Date, pen.Pension_Scheme
			FROM inserted i
			INNER JOIN Contract_Templates ct ON ct.Contract = i.Contract
			INNER JOIN Contract_Pension_Schemes pen ON pen.ID_215 = ct.ID AND (pen.End_Date >= GETDATE() OR End_Date IS NULL);

	INSERT Post_Working_Patterns (ID_219, Effective_Date, Regional_ID, Absence_In, Day_Pattern, Sunday_Hours_AM, Sunday_Hours_PM, Monday_Hours_AM, Monday_Hours_PM
									, Tuesday_Hours_AM, Tuesday_Hours_PM, Wednesday_Hours_AM, Wednesday_Hours_PM, Thursday_Hours_AM, Thursday_Hours_PM
									, Friday_Hours_AM, Friday_Hours_PM, Saturday_Hours_AM, Saturday_Hours_PM)
		SELECT i.ID, i.Effective_Date, wp.Regional_ID, wp.Absence_In, wp.Day_Pattern, wp.Sunday_Hours_AM, wp.Sunday_Hours_PM, wp.Monday_Hours_AM,wp.Monday_Hours_PM
									, wp.Tuesday_Hours_AM, wp.Tuesday_Hours_PM, wp.Wednesday_Hours_AM, wp.Wednesday_Hours_PM, wp.Thursday_Hours_AM, wp.Thursday_Hours_PM
									, wp.Friday_Hours_AM, wp.Friday_Hours_PM, wp.Saturday_Hours_AM, wp.Saturday_Hours_PM
			FROM inserted i
			INNER JOIN Contract_Templates ct ON ct.Contract = i.Contract
			INNER JOIN Contract_Working_Patterns wp ON wp.ID_215 = ct.ID AND (wp.End_Date >= GETDATE() OR End_Date IS NULL);');

GO

INSERT ASRSysTableTriggers (TriggerID, TableID, Name, CodePosition, IsSystem, Content) VALUES (15, 228, 'Transfer child information based on contract type', 1, 1, ' 	-- Insert leave schemes based on contract type
	INSERT Post_Leave_Schemes (ID_219, Effective_Date, End_Date, Leave_Scheme, Notes)
		SELECT i.ID_219, i.Effective_Date, chs.End_Date, chs.Leave_Scheme, chs.Notes
			FROM Contract_Leave_Schemes chs
			INNER JOIN Contract_Templates ct ON ct.ID = chs.ID_215
			INNER JOIN inserted i ON i.Contract = ct.Contract
			INNER JOIN deleted d ON i.id = d.id
			WHERE i.Contract <> d.Contract OR @startingtrigger = 1 AND chs.Effective_Date <= GETDATE() AND (chs.End_Date IS NULL OR chs.End_Date >= GETDATE());

	-- Insert OMP Schemes based on contract type
	INSERT Post_OMP_Schemes (ID_219, Effective_Date, End_Date, OMP_Scheme, Notes, Description)
		SELECT i.ID_219, i.Effective_Date, cms.End_Date, cms.OMP_Scheme, cms.Notes, cms.Description
			FROM Contract_OMP_Schemes cms
			INNER JOIN Contract_Templates ct ON ct.ID = cms.ID_215
			INNER JOIN inserted i ON i.Contract = ct.Contract
			INNER JOIN deleted d ON i.id = d.id
			WHERE i.Contract <> d.Contract OR @startingtrigger = 1 AND cms.Effective_Date <= GETDATE() AND (cms.End_Date IS NULL OR cms.End_Date >= GETDATE());

	-- Insert OSP Schemes based on contract type
	INSERT Post_OSP_Schemes (ID_219, Effective_Date, End_Date, OSP_Scheme, Notes, Description)
		SELECT i.ID_219, i.Effective_Date, css.End_Date, css.OSP_Scheme, css.Notes, css.Description
			FROM Contract_OSP_Schemes css
			INNER JOIN Contract_Templates ct ON ct.ID = css.ID_215
			INNER JOIN inserted i ON i.Contract = ct.Contract
			INNER JOIN deleted d ON i.id = d.id
			WHERE i.Contract <> d.Contract OR @startingtrigger = 1 AND css.Effective_Date <= GETDATE() AND (css.End_Date IS NULL OR css.End_Date >= GETDATE());

	-- Insert OSP Schemes based on contract type
	INSERT Post_Pension_Schemes (ID_219, Effective_Date, End_Date, Pension_Scheme, Scheme_Number, Notes)
		SELECT i.ID_219, i.Effective_Date, cps.End_Date, cps.Pension_Scheme, cps.Scheme_Number, cps.Notes
			FROM Contract_Pension_Schemes cps
			INNER JOIN Contract_Templates ct ON ct.ID = cps.ID_215
			INNER JOIN inserted i ON i.Contract = ct.Contract
			INNER JOIN deleted d ON i.id = d.id
			WHERE i.Contract <> d.Contract OR @startingtrigger = 1 AND cps.Effective_Date <= GETDATE() AND (cps.End_Date IS NULL OR cps.End_Date >= GETDATE());

	-- Insert Working Pattern Schemes based on contract type
	INSERT Post_Working_Patterns (ID_219, Effective_Date, Regional_ID, Absence_In, Sunday_Hours_AM, Sunday_Hours_PM, Monday_Hours_AM, Monday_Hours_PM
									, Tuesday_Hours_AM, Tuesday_Hours_PM, Wednesday_Hours_AM, Wednesday_Hours_PM, Thursday_Hours_AM, Thursday_Hours_PM
									, Friday_Hours_AM, Friday_Hours_PM, Saturday_Hours_AM, Saturday_Hours_PM)
		SELECT i.ID_219, i.Effective_Date, cws.Regional_ID, cws.Absence_In, cws.Sunday_Hours_AM, cws.Sunday_Hours_PM, cws.Monday_Hours_AM,cws.Monday_Hours_PM
									, cws.Tuesday_Hours_AM, cws.Tuesday_Hours_PM, cws.Wednesday_Hours_AM, cws.Wednesday_Hours_PM, cws.Thursday_Hours_AM, cws.Thursday_Hours_PM
									, cws.Friday_Hours_AM, cws.Friday_Hours_PM, cws.Saturday_Hours_AM, cws.Saturday_Hours_PM
			FROM Contract_Working_Patterns cws
			INNER JOIN Contract_Templates ct ON ct.ID = cws.ID_215
			INNER JOIN inserted i ON i.Contract = ct.Contract
			INNER JOIN deleted d ON i.id = d.id
			WHERE i.Contract <> d.Contract OR @startingtrigger = 1 AND cws.Effective_Date <= GETDATE() AND (cws.End_Date IS NULL OR cws.End_Date >= GETDATE());');
GO

INSERT ASRSysTableTriggers (TriggerID, TableID, Name, CodePosition, IsSystem, Content) VALUES (16, 252, 'Update appointment duration', 1, 1, ' 	-- Update duration of appointment absence
	IF @startingtriggertable = 252
        UPDATE [dbo].[tbuser_Appointment_Absence] SET [updflag] = 1 WHERE [dbo].[tbuser_Appointment_Absence].[id] IN (SELECT DISTINCT [id_251] FROM inserted)
');
GO

