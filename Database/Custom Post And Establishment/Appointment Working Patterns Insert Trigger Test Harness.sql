
DELETE FROM Appointments
 delete from Appointment_Working_Patterns
 delete from Working_Patterns
 

--SELECT TOP 1 @newID = ID FROM Appointment_Working_Patterns ORDER BY ID DESC

DECLARE @newID integer;

INSERT Appointments (Appointment_Start_Date, ID_219, ID_1) VALUES (GETDATE()-50, 162, 116)

--SELECT TOP 1 @newID = ID FROM Appointment_Working_Patterns ORDER BY ID DESC
--UPDATE Appointment_Working_Patterns SET sunday_hours_AM = 0, sunday_hours_PM = 0, Monday_Hours_AM=0, Monday_Hours_PM=0 where id = @newID

SELECT TOP 1 @newID = ID FROM Appointments ORDER BY ID DESC
DELETE FROM Appointment_Working_Patterns WHERE ID_3 = @newID
INSERT Appointment_Working_Patterns (Effective_Date, End_Date, ID_3, Monday_Hours_AM, Tuesday_Hours_AM, Wednesday_Hours_AM) VALUES ('2015-01-01', '2015-01-31',  @newID, 1, 1, 1)
INSERT Appointment_Working_Patterns (Effective_Date, End_Date, ID_3, Wednesday_Hours_AM, Friday_Hours_AM) VALUES ('2015-01-05','2015-02-17',  @newID, 1, 1)
INSERT Appointment_Working_Patterns (Effective_Date, End_Date, ID_3, Monday_Hours_AM, Saturday_Hours_AM) VALUES ('2015-01-01','2015-01-14',  @newID, 1, 1)
INSERT Appointment_Working_Patterns (Effective_Date, End_Date, ID_3, Thursday_Hours_AM, Saturday_Hours_AM) VALUES ('2015-02-18',NULL,  @newID, 2.25, 1.5)
INSERT Appointment_Working_Patterns (Effective_Date, End_Date, ID_3, Sunday_Hours_AM, Sunday_Hours_PM) VALUES ('2015-01-09','2015-01-09',  @newID, 18, 18)

--INSERT Appointments (Appointment_Start_Date, ID_219, ID_1) VALUES (GETDATE() -25, 163, 116)
--SELECT TOP 1 @newID = ID FROM Appointments ORDER BY ID DESC

--SELECT TOP 1 @newID = ID FROM Appointment_Working_Patterns ORDER BY ID DESC
--INSERT Appointment_Working_Patterns (Effective_Date, End_Date, ID_3, Monday_Hours_AM, Tuesday_Hours_AM, Wednesday_Hours_AM) VALUES ('2015-01-01', '2015-01-31',  @newID, 1, 1, 1)
 
INSERT Appointments (Appointment_Start_Date, ID_219, ID_1) VALUES (GETDATE(), 164, 116)
SELECT TOP 1 @newID = ID FROM Appointments ORDER BY ID DESC
DELETE FROM Appointment_Working_Patterns WHERE ID_3 = @newID
INSERT Appointment_Working_Patterns (Effective_Date, End_Date, ID_3, Tuesday_Hours_AM, Friday_Hours_AM) VALUES ('2015-01-01', '2015-01-31',  @newID, 1, 1)


--SELECT TOP 1 @newID = ID FROM Appointment_Working_Patterns ORDER BY ID DESC

--select * from ASRSysTables order by tablename

--select ID_1, ID_219, * from Appointments WHERE ID = 876
SELECT * FROM Appointments
select id, Effective_Date, End_Date, * from Appointment_Working_Patterns order by 2 -- where ID_3 = 876 ORDER BY 1
select * from Working_Patterns order by 1
--select * from Appointment_Working_Patterns

--DECLARE @persID integer;
--DECLARE @changeDates TABLE(effectiveDate datetime, ID integer, PersID integer);


--INSERT @changeDates 
--	SELECT awp.Effective_Date, awp.ID , a.ID_1
--		FROM Appointment_Working_Patterns awp
--		INNER JOIN Appointments a ON a.ID = awp.ID_3	
--		WHERE ID_3 = 876

--INSERT @changeDates 
--	SELECT awp.End_Date + 1, awp.ID, a.ID_1
--		FROM Appointment_Working_Patterns awp
--		INNER JOIN Appointments a ON a.ID = awp.ID_3	
--		WHERE ID_3 = 876


--DECLARE @merged TABLE (effective_date datetime, sunHoursAM numeric(4,2), sunHoursPM numeric(4,2)
--						, MonHoursAM numeric(4,2), MonHoursPM numeric(4,2)
--						, TuesHoursAM numeric(4,2), TuesHoursPM numeric(4,2)
--						, WedHoursAM numeric(4,2), WedHoursPM numeric(4,2)
--						, ThursHoursAM numeric(4,2), ThursHoursPM numeric(4,2)
--						, FriHoursAM numeric(4,2), FriHoursPM numeric(4,2)
--						, SatHoursAM numeric(4,2), SatHoursPM numeric(4,2))

--	DECLARE @cursRollupWorkingPatterns cursor,
--		@effectiveDate datetime;

--    SET @cursRollupWorkingPatterns = CURSOR LOCAL FAST_FORWARD READ_ONLY FOR 
--		SELECT DISTINCT EffectiveDate FROM @changeDates
--			ORDER BY EffectiveDate;

--	OPEN @cursRollupWorkingPatterns
--	FETCH NEXT FROM @cursRollupWorkingPatterns INTO @effectiveDate
--    WHILE (@@fetch_status = 0)
--	BEGIN

--		INSERT @merged
--			SELECT @effectiveDate
--			, ISNULL(SUM(wp.Sunday_Hours_AM),0), ISNULL(SUM(wp.Sunday_Hours_PM),0)
--			, ISNULL(SUM(wp.Monday_Hours_AM),0), ISNULL(SUM(wp.Monday_Hours_PM),0)
--			, ISNULL(SUM(wp.Tuesday_Hours_AM),0), ISNULL(SUM(wp.Tuesday_Hours_PM),0)
--			, ISNULL(SUM(wp.Wednesday_Hours_AM),0), ISNULL(SUM(wp.Wednesday_Hours_PM),0)
--			, ISNULL(SUM(wp.Thursday_Hours_AM),0), ISNULL(SUM(wp.Thursday_Hours_PM),0)
--			, ISNULL(SUM(wp.Friday_Hours_AM),0), ISNULL(SUM(wp.Friday_Hours_PM),0)
--			, ISNULL(SUM(wp.Saturday_Hours_AM),0), ISNULL(SUM(wp.Saturday_Hours_PM),0)

--		FROM Appointment_Working_Patterns wp
--			WHERE @effectiveDate >= wp.Effective_Date AND @effectiveDate <= wp.End_Date;

--		FETCH NEXT FROM @cursRollupWorkingPatterns INTO @effectiveDate
--	END
--	CLOSE @cursRollupWorkingPatterns;
--    DEALLOCATE @cursRollupWorkingPatterns;

--	--select * from Working_Patterns
--	--INSERT Working_Patterns (Effective_Date, Sunday_Hours, Saturday_Hours, Monday_Hours, Tuesday_Hours, Wednesday_Hours, Thursday_Hours, Friday_Hours, Saturday_Hours, Working_Pattern)
--		SELECT effective_date, sunHoursAM + sunHoursPM, MonHoursAM + MonHoursPM, TuesHoursAM + TuesHoursPM
--			, WedHoursAM + WedHoursPM, ThursHoursAM + ThursHoursPM, FriHoursAM + FriHoursPM, SatHoursAM + SatHoursPM, 
--			CASE WHEN sunHoursAM > 0 THEN 'S' ELSE ' ' END + 
--			CASE WHEN sunHoursPM > 0 THEN 'S' ELSE ' ' END +
--			CASE WHEN MonHoursAM > 0 THEN 'M' ELSE ' ' END +
--			CASE WHEN MonHoursPM > 0 THEN 'M' ELSE ' ' END +
--			CASE WHEN TuesHoursAM > 0 THEN 'T' ELSE ' ' END +
--			CASE WHEN TuesHoursPM > 0 THEN 'T' ELSE ' ' END +
--			CASE WHEN WedHoursAM > 0 THEN 'W' ELSE ' ' END +
--			CASE WHEN WedHoursPM > 0 THEN 'W' ELSE ' ' END +
--			CASE WHEN ThursHoursAM > 0 THEN 'T' ELSE ' ' END +
--			CASE WHEN ThursHoursPM > 0 THEN 'T' ELSE ' ' END +
--			CASE WHEN FriHoursAM > 0 THEN 'F' ELSE ' ' END +
--			CASE WHEN FriHoursPM > 0 THEN 'F' ELSE ' ' END +
--			CASE WHEN SatHoursAM > 0 THEN 'S' ELSE ' ' END +
--			CASE WHEN SatHoursPM > 0 THEN 'S' ELSE ' ' END
--		FROM @merged;

