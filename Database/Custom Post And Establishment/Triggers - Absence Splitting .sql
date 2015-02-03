
/* Required Tables

249 - Appointment_Working_Patterns
250	- Absence_Entry
251	- Appointment_Absence_Entry
252	- Absence_Breakdown

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
UPDATE ASRSysTables SET deletetriggerdisabled = 1, inserttriggerdisabled = 1 WHERE tableid = 252;


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
	INSERT Appointment_Absence_Staging (ID_3, Start_Date, Start_Session, End_Date, End_Session, Absence_Type, Reason, Duration, Absence_In)
		SELECT a.ID, ae.Start_Date, ae.Start_Session, ae.End_Date, ae.End_Session
				, ae.Absence_Type, ae.Reason
				, dbo.udfcustom_AbsenceDurationForAppointment(ae.Start_Date, ae.Start_Session, ae.End_Date, ae.End_Session, a.ID)
				, a.Absence_In
		FROM inserted ae
		INNER JOIN Appointments a ON ae.ID_1 = a.ID_1 AND (a.post_ID = ae.post_ID OR ae.post_ID IS NULL)
		WHERE ae.Start_Date <= ISNULL(Appointment_End_Date, convert(datetime, ''9999-12-31''))
			AND ae.End_Date >= ISNULL(Appointment_Start_Date, convert(datetime, ''1899-12-31''));

')
GO

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

	INSERT Appointment_Holiday_Schemes(ID_3, Effective_Date, Holiday_Scheme)
		SELECT i.ID, i.Appointment_Start_Date, phs.Holiday_Scheme
			FROM inserted i
			INNER JOIN Post_Holiday_Schemes phs ON phs.ID_219 = i.ID_219;

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
		USING (SELECT PersID, effective_date, sunHoursAM + sunHoursPM AS sunHours, MonHoursAM + MonHoursPM AS MonHours, TuesHoursAM + TuesHoursPM AS TuesHours
			, WedHoursAM + WedHoursPM AS WedHours, ThursHoursAM + ThursHoursPM AS ThursHours, FriHoursAM + FriHoursPM AS FriHours, SatHoursAM + SatHoursPM AS SatHours, 
			CASE WHEN sunHoursAM > 0 THEN ''S'' ELSE '' '' END + 
			CASE WHEN sunHoursPM > 0 THEN ''S'' ELSE '' '' END +
			CASE WHEN MonHoursAM > 0 THEN ''M'' ELSE '' '' END +
			CASE WHEN MonHoursPM > 0 THEN ''M'' ELSE '' '' END +
			CASE WHEN TuesHoursAM > 0 THEN ''T'' ELSE '' '' END +
			CASE WHEN TuesHoursPM > 0 THEN ''T'' ELSE '' '' END +
			CASE WHEN WedHoursAM > 0 THEN ''W'' ELSE '' '' END +
			CASE WHEN WedHoursPM > 0 THEN ''W'' ELSE '' '' END +
			CASE WHEN ThursHoursAM > 0 THEN ''T'' ELSE '' '' END +
			CASE WHEN ThursHoursPM > 0 THEN ''T'' ELSE '' '' END +
			CASE WHEN FriHoursAM > 0 THEN ''F'' ELSE '' '' END +
			CASE WHEN FriHoursPM > 0 THEN ''F'' ELSE '' '' END +
			CASE WHEN SatHoursAM > 0 THEN ''S'' ELSE '' '' END +
			CASE WHEN SatHoursPM > 0 THEN ''S'' ELSE '' '' END AS workpatt
		FROM @merged)
	AS awp ON (wp.ID_1 = awp.PersID AND wp.Effective_Date = awp.effective_date) 
	WHEN NOT MATCHED BY TARGET
		THEN INSERT(ID_1, Effective_Date, Sunday_Hours, Monday_Hours, Tuesday_Hours, Wednesday_Hours, Thursday_Hours, Friday_Hours, Saturday_Hours, Working_Pattern) 
			VALUES(awp.persID, awp.Effective_Date, awp.SunHours, awp.MonHours, awp.TuesHours, awp.WedHours, awp.ThursHours, awp.FriHours, awp.SatHours, awp.workpatt)
	WHEN MATCHED 
		THEN UPDATE SET Effective_Date = awp.Effective_Date, Sunday_Hours = awp.SunHours, Monday_Hours = awp.MonHours, Tuesday_Hours = awp.TuesHours, Wednesday_Hours = awp.WedHours
			, Thursday_Hours = awp.ThursHours, Friday_Hours = awp.FriHours, Saturday_Hours = awp.SatHours, Working_Pattern = awp.workpatt
	WHEN NOT MATCHED BY SOURCE AND wp.ID_1 IN(SELECT PersID FROM @employees)
	    THEN DELETE ;');

GO

INSERT ASRSysTableTriggers (TriggerID, TableID, Name, CodePosition, IsSystem, Content) VALUES (7, 219, 'Transfer working pattern from appointment', 1, 1, '    INSERT Post_Holiday_Schemes (ID_219, Effective_Date, Holiday_Scheme)
		SELECT i.ID, i.Effective_Date, chs.Holiday_Scheme
			FROM inserted i
			INNER JOIN Contract_Templates ct ON ct.Contract = i.Contract
			INNER JOIN Contract_Holiday_Schemes chs ON chs.ID_215 = ct.ID AND (chs.End_Date >= GETDATE() OR End_Date IS NULL);

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
