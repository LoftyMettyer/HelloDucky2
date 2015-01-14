/* Required Tables

249 - Appointment_Working_Patterns
250	- Absence_Entry
251	- Appointment_Absence_Entry
252	- Absence_Breakdown

*/

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

INSERT ASRSysTableTriggers (TriggerID, TableID, Name, CodePosition, IsSystem, Content) VALUES (1, 251, 'Update Absence Breakdown', 1, 1, '		DELETE [dbo].[tbuser_Absence_Breakdown] WHERE [id_251] IN (SELECT DISTINCT [id] FROM deleted);

	DECLARE @changeDates TABLE(Effective_Date datetime, AppointmentID integer, PersID integer);

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

	INSERT Absence_Breakdown([source], ID_251, Post_ID, [Type], Payroll_Type_Code, Reason, Payroll_Reason_Code, Absence_In, Duration, Absence_Date, [Session]
		, Day_Pattern_AM, Day_Pattern_PM, Hour_Pattern_AM, Hour_Pattern_PM, Staff_Number, Payroll_Company_Code)	
		SELECT ''post'', i.ID, ap.Post_ID, i.Absence_Type, ISNULL(at.Payroll_Code, ''''), i.Reason, ISNULL(ar.Payroll_Code, ''''), ap.Absence_In
			, dbo.udfsysDurationFromPattern(ap.Absence_In, dr.IndividualDate, dr.SessionType, wp.Sunday_Hours_AM, wp.Monday_Hours_AM, wp.Tuesday_Hours_AM, wp.Wednesday_Hours_AM, wp.Thursday_Hours_AM, wp.Friday_Hours_AM, wp.Saturday_Hours_AM, wp.Sunday_Hours_PM, wp.Monday_Hours_PM, wp.Tuesday_Hours_PM, wp.Wednesday_Hours_PM, wp.Thursday_Hours_PM, wp.Friday_Hours_PM, wp.Saturday_Hours_PM)
			, dr.IndividualDate, dr.SessionType
			, wp.Day_Pattern_AM, wp.Day_Pattern_PM, wp.Hour_Pattern_AM, wp.Hour_Pattern_PM
			, pr.Staff_Number, pr.Payroll_Company_Code
		FROM inserted i
			CROSS APPLY [dbo].[udfsysDateRangeToTable] (''d'', i.Start_Date, i.Start_Session,  i.End_Date, i.End_Session) dr
			INNER JOIN Appointments ap ON ap.ID = i.ID_3
			INNER JOIN @working_patterns wp ON wp.AppointmentID = ap.ID
			INNER JOIN Personnel_Records pr ON pr.ID = ap.ID_1
			LEFT JOIN Absence_Type_Table at ON at.Absence_Type = i.Absence_Type
			LEFT JOIN Absence_Reason_Table ar ON ar.Reason = i.Reason
		WHERE wp.Effective_Date <= dr.IndividualDate AND (wp.End_Date >= dr.IndividualDate OR wp.End_Date IS NULL);')
GO

INSERT ASRSysTableTriggers (TriggerID, TableID, Name, CodePosition, IsSystem, Content) VALUES (2, 250, 'Update Absence Breakdown', 1, 1, '    DELETE [dbo].[tbuser_Absence_Breakdown] WHERE [id_250] IN (SELECT DISTINCT [id] FROM deleted);

	DECLARE @persID integer;
	DECLARE @changeDates TABLE(Effective_Date datetime, AppointmentID integer, PersID integer);

	INSERT @changeDates 
		SELECT awp.Effective_Date, a.ID, e.ID_1 FROM Appointments a
			INNER JOIN inserted e ON e.ID_1 = a.ID_1
			INNER JOIN Appointment_Working_Patterns awp ON awp.ID_3 = a.ID
		WHERE awp.Effective_Date <= e.Start_Date
			AND (a.Appointment_End_Date >= e.Start_Date OR a.Appointment_End_Date IS NULL)
	  UNION
		SELECT awp.End_Date + 1, a.ID, e.ID_1 FROM Appointments a
			INNER JOIN inserted e ON e.ID_1 = a.ID_1
			INNER JOIN Appointment_Working_Patterns awp ON awp.ID_3 = a.ID
		WHERE awp.End_Date IS NOT NULL
			AND (a.Appointment_End_Date >= e.Start_Date OR a.Appointment_End_Date IS NULL);

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
		SELECT cd.PersID, cd.Effective_Date, NULL, cd.AppointmentID
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
			GROUP BY  cd.AppointmentID, cd.PersID, cd.Effective_Date;

	UPDATE t 
		SET End_Date = (SELECT top 1 m.Effective_Date - 1 FROM @working_patterns m WHERE m.Effective_Date > t.Effective_Date AND m.AppointmentID = t.AppointmentID ORDER BY m.Effective_Date),
			Day_Pattern_AM = dbo.udfsysPatternFromHours (''Days'', Sunday_Hours_AM, Monday_Hours_AM, Tuesday_Hours_AM, Wednesday_Hours_AM, Thursday_Hours_AM, Friday_Hours_AM, Saturday_Hours_AM),
			Day_Pattern_PM = dbo.udfsysPatternFromHours (''Days'', Sunday_Hours_PM, Monday_Hours_PM, Tuesday_Hours_PM, Wednesday_Hours_PM, Thursday_Hours_PM, Friday_Hours_PM, Saturday_Hours_PM),
			Hour_Pattern_AM = dbo.udfsysPatternFromHours (''Hours'', Sunday_Hours_AM, Monday_Hours_AM, Tuesday_Hours_AM, Wednesday_Hours_AM, Thursday_Hours_AM, Friday_Hours_AM, Saturday_Hours_AM),
			Hour_Pattern_PM = dbo.udfsysPatternFromHours (''Hours'', Sunday_Hours_PM, Monday_Hours_PM, Tuesday_Hours_PM, Wednesday_Hours_PM, Thursday_Hours_PM, Friday_Hours_PM, Saturday_Hours_PM)
	FROM @working_patterns t

	   -- Update Absence Breakdown
    DELETE [dbo].[tbuser_Absence_Breakdown] WHERE [id_250] IN (SELECT DISTINCT [id] FROM deleted);

	INSERT Absence_Breakdown([source], ID_250, Post_ID, [Type], Payroll_Type_Code, Reason, Payroll_Reason_Code, Absence_In, Duration, Absence_Date, [Session]
		, Day_Pattern_AM, Day_Pattern_PM, Hour_Pattern_AM, Hour_Pattern_PM, Staff_Number, Payroll_Company_Code)	
		SELECT ''pers'', i.ID, ap.Post_ID, i.Absence_Type, ISNULL(at.Payroll_Code, ''''), i.Reason, ISNULL(ar.Payroll_Code, ''''), ap.Absence_In,
			dbo.udfsysDurationFromPattern(ap.Absence_In, dr.IndividualDate, dr.SessionType, wp.Sunday_Hours_AM, wp.Monday_Hours_AM, wp.Tuesday_Hours_AM
			, wp.Wednesday_Hours_AM, wp.Thursday_Hours_AM, wp.Friday_Hours_AM, wp.Saturday_Hours_AM, wp.Sunday_Hours_PM, wp.Monday_Hours_PM, wp.Tuesday_Hours_PM
			, wp.Wednesday_Hours_PM, wp.Thursday_Hours_PM, wp.Friday_Hours_PM, wp.Saturday_Hours_PM)
		AS Duration
			, dr.IndividualDate, dr.SessionType
			, wp.Day_Pattern_AM, wp.Day_Pattern_PM, wp.Hour_Pattern_AM, wp.Hour_Pattern_PM
			, pr.Staff_Number, pr.Payroll_Company_Code
		FROM inserted i
			CROSS APPLY [dbo].[udfsysDateRangeToTable] (''d'', i.Start_Date, i.Start_Session,  i.End_Date, i.End_Session) dr
			INNER JOIN Appointments ap ON ap.ID_1 = i.ID_1
			INNER JOIN @working_patterns wp ON wp.AppointmentID = ap.ID
			INNER JOIN Personnel_Records pr ON pr.ID = i.ID_1
			LEFT JOIN Absence_Type_Table at ON at.Absence_Type = i.Absence_Type
			LEFT JOIN Absence_Reason_Table ar ON ar.Reason = i.Reason
		WHERE wp.Effective_Date <= dr.IndividualDate AND (wp.End_Date >= dr.IndividualDate OR wp.End_Date IS NULL);')
GO

INSERT ASRSysTableTriggers (TriggerID, TableID, Name, CodePosition, IsSystem, Content) VALUES (3, 252, 'Absence Rollup', 1, 1, '	DECLARE @thisAbsenceID	integer,
			@startDate	datetime,
			@endDate	datetime;

	INSERT Absence(Absence_Type, Payroll_Code, Reason, Payroll_Reason, Start_Date, Start_Session, End_Date, End_Session, Duration, ID_1, Absence_In)
		SELECT DISTINCT ab.Type, ab.Payroll_Type_Code, ab.Reason, ab.Payroll_Reason_Code
			, m.startdate
			, (SELECT DISTINCT CASE WHEN [Session] = ''Day'' THEN ''AM'' ELSE [Session] END FROM inserted WHERE (ID_250 = ab.ID_250 OR ID_251 = ab.ID_251) AND Absence_Date = m.startdate)
			, m.enddate
			, (SELECT DISTINCT CASE WHEN [Session] = ''Day'' THEN ''PM'' ELSE [Session] END FROM inserted WHERE (ID_250 = ab.ID_250 OR ID_251 = ab.ID_251) AND Absence_Date = m.enddate)
			, m.Duration
			, CASE WHEN  ab.[source] = ''pers'' THEN ae.ID_1 ELSE a.ID_1 END
			, ab.Absence_In
		FROM inserted ab
			LEFT JOIN Absence_Entry ae ON ae.ID = ab.ID_250
			LEFT JOIN Appointment_Absence_Entry aae ON aae.ID = ab.ID_251
			LEFT JOIN Appointments a ON a.ID = aae.ID_3
		CROSS APPLY (
			SELECT MIN(range.Absence_Date) AS startdate, MAX(range.Absence_Date) AS enddate, SUM(Duration) AS Duration
			FROM inserted range
			WHERE ab.ID_250 = range.ID_250 OR ab.ID_251 = range.ID_251) m;


	SELECT @thisAbsenceID = MAX(ID) FROM Absence;

	UPDATE [tbuser_Absence_Breakdown]
		SET id_2 = @thisAbsenceID
	FROM [inserted] base WHERE base.[id] = [dbo].[tbuser_Absence_Breakdown].[id];');

GO

INSERT ASRSysTableTriggers (TriggerID, TableID, Name, CodePosition, IsSystem, Content) VALUES (4, 252, 'Absence Deletion', 2, 1, '	DECLARE @AbsenceID	integer,
			@startDate	datetime,
			@endDate	datetime;

    DELETE [dbo].[tbuser_Absence] WHERE [id] IN (SELECT DISTINCT [id_2] FROM deleted);

	WITH base AS (SELECT * FROM dbo.[tbuser_Absence_Breakdown]
        WHERE [id] IN (SELECT DISTINCT [id] FROM deleted))
        DELETE FROM base;

END')
GO

INSERT ASRSysTableTriggers (TriggerID, TableID, Name, CodePosition, IsSystem, Content) VALUES (5, 3, 'Populate from Post Template', 1, 1, '    INSERT Appointment_Allowances(ID_3, Effective_Date, Type, Frequency, Amount, Currency)
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
