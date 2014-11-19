
-- delete from Appointment_Working_Patterns
-- delete from Working_Patterns

DECLARE @newID integer;

--INSERT Appointment_Working_Patterns (Effective_Date, End_Date, ID_3, Sunday_Hours_AM, Sunday_Hours_PM) VALUES (GETDATE()-3, GETDATE()+2,  876, 9.2, 0.8)
INSERT Appointment_Working_Patterns (Effective_Date, End_Date, ID_3, Sunday_Hours_AM, Sunday_Hours_PM) VALUES (GETDATE()-3, GETDATE()+4,  876, 9.2, 0.7)

SELECT TOP 1 @newID = ID FROM Appointment_Working_Patterns ORDER BY ID DESC

--select * from ASRSysTables order by tablename

--select ID_1, ID_219, * from Appointments WHERE ID = 876
--select * from Appointment_Working_Patterns where ID_3 = 876
select * from Working_Patterns
--select * from Appointment_Working_Patterns

--update Appointment_Working_Patterns set Sunday_Hours_AM = 1.2, Sunday_Hours_PM = 1.1 where ID in (36, 37)