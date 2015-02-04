DECLARE @persID integer

--select id_1 from Appointments order by 1
select @persID = id from Personnel_Records where surname = 'Clipper'

--SELECT id,* FROM Appointments WHERE ID_1 = 116
--
--SELECT id,* FROM Appointment_Working_Patterns WHERE ID_3 = 876
--delete from Appointment_Working_Patterns where ID in (130)
--UPDATE Appointment_Working_Patterns SET Effective_Date = '2014-11-17' WHERE ID = 126
--UPDATE Appointment_Working_Patterns SET friDAY_HOURS_aM = 1, fridaY_hOURS_pM = 2  WHERE ID = 126

DELETE FROM ASRSysAccordTransactions

DELETE FROM Absence_Entry
DELETE FROM Appointment_Absence_Staging
DELETE FROM Appointment_Absence
DELETE FROM Absence_Breakdown
DELETE FROM Absence

    --/* This sets all of the flags prior to updating date dependant columns */
    --DELETE FROM ASRSYSSystemSettings WHERE [Section] = 'database' and [SettingKey] = 'updatingdatedependantcolumns'

    --INSERT ASRSYSSystemSettings([Section],[SettingKey],[SettingValue])
    --VALUES('database','updatingdatedependantcolumns',1)

print 'iNSERT'
--delete from Appointment_Working_Patterns
--INSERT Appointment_Absence_Entry (ID_3, start_date, start_session, end_date, end_session, Reason, absence_type) values (825, getdate()-1000, 'AM', getdate()-995, 'PM', 'multiple post day off', 'HOLS')
--INSERT Appointment_Absence_Entry (ID_3, start_date, start_session, end_date, end_session, Reason, absence_type) values (825, getdate()-1, 'AM', getdate()-1, 'PM', 'just a single day off', 'MAT')
--INSERT Appointment_Absence_Entry (ID_3, start_date, start_session, end_date, end_session, Reason, absence_type) values (825, getdate(), 'AM', getdate(), 'AM', 'morning', 'HOLS')
--INSERT Appointment_Absence_Entry (ID_3, start_date, start_session, end_date, end_session, Reason, absence_type) values (825, getdate()+1, 'PM', getdate()+1, 'PM', 'afternoon', 'JURY')

--SELECT id,* FROM Appointments where ID_1 = @persID 
--SELECT id,* FROM Appointment_Working_Patterns order by Effective_Date --where id_3 in (835, 876)
--select Effective_Date, *  from Appointment_Working_Patterns where id_3 = 1108

--INSERT Appointment_Absence_Entry (ID_3, start_date, start_session, end_date, end_session, Reason, absence_type) values (1108, '2014-01-01', 'PM', '2014-04-30', 'PM', 'malaria', 'Holiday2')

--INSERT Absence_Entry (ID_1, start_date, end_date, Reason, absence_type) values (116, getdate(), getdate()+2, 'SICK', 'felt bad')
--INSERT Absence_Entry (ID_1, start_date, end_date, Reason, absence_type) values (116, getdate()-1, getdate()-1, 'SICK', 'felt yucky')
--select * from absence_type_table
--INSERT Absence_Entry (ID_1, start_date, end_date, absence_type, Reason) values (@persID, getdate(), getdate()+10, 'HOLS', 'bit of jolly')
INSERT Absence_Entry (ID_1, start_date, Start_Session, end_date, end_session, absence_type, Reason) values (@persID, getdate()-7, 'PM', NULL, '', 'Sickness', 'felt bad')
--INSERT Absence_Entry (ID_1, start_date, end_date, absence_type, Reason) values (@persID, getdate(), getdate()+10, 'Maternity Leave', 'mummy!')
--INSERT Absence_Entry (ID_1, start_date, end_date, absence_type, Reason) values (@persID, getdate(), getdate()+10, 'Holiday', 'Lots of sun')


--INSERT Absence_Entry (ID_1, start_date, end_date, Reason, absence_type) values (117, getdate()-10, getdate()-10, 'HOLS', 'bit of jolly')
--INSERT Absence_Entry (ID_1, start_date, end_date, Reason, absence_type) values (116,'2016-01-01', '2016-01-07', 'HOLS', 'jolly16')
--INSERT Absence_Entry (ID_1, start_date, end_date, Reason, absence_type) values (116,'2015-01-01', '2015-01-07', 'HOLS', 'bit of jolly')

--INSERT Absence_Entry (ID_1, start_date, start_session, end_date, end_session, absence_type, Reason, Absence_In) values (116, getdate()-5, 'PM', getdate()-5, 'AM', 'SICK', 'Jody felt bad', 'Hours')
--INSERT Absence_Entry (ID_1, start_date, start_session, end_date, end_session, absence_type, Reason) values (116, getdate()-10, 'AM', getdate()+10, 'PM', 'sickness', 'bad back')

PRINT 'UPDATE'
--UPDATE Absence_Entry SET start_date = start_date - 1
--UPDATE Appointment_Absence_Entry SET start_date = end_date - 1
-- DELETE FROM Absence_Entry WHERE ID = 521
-- DELETE FROM Appointment_Absence_Entry --WHERE ID = 86
--SELECT * FROM Absence_Entry

--SELECT * FROM Appointment_Absence_Entry
--SELECT * FROM Absence_Entry

--select * from Working_Patterns
--SELECT * FROM Absence_Breakdown ORDER BY Absence_Date ASC
--SELECT id, Start_date, Start_Session, End_Date, End_Session, Reason, * FROM Absence
--SELECT ID,Start_Date, Start_Session, End_Date, End_Session, Reason, Absence_Type, Duration_Days, Duration_Hours FROM Absence

--SELECT * FROM ASRSysAccordTransactions
--SELECT Start_Date, End_Date, Duration_Hours, Duration_Days FROM Absence

--select * from ASRSysTables order by tablename
--select * from ASRSysAccordTransferTypes order by ASRBaseTableID
--SELECT * FROM Appointments where ID =  876

--WHERE wp.Effective_Date <= dr.IndividualDate AND (wp.End_Date >= dr.IndividualDate OR wp.End_Date IS NULL);

--update Appointments set Absence_In = 'Days' where id = 1108

--select * from Absence_Type_Table

--UPDATE Appointment_Absence_Staging SET Authorised = 1 WHERE Absence_Type = 'HOLS'
--UPDATE Appointment_Absence_Staging SET Authorised = 1 WHERE ID = 270 -- WHERE Absence_Type = 'HOLS'
--update Appointment_Absence set End_Date = getdate() + 10 where id = 736

SELECT Duration, * FROM Appointment_Absence_Staging
SELECT Duration, * FROM Appointment_Absence
SELECT * FROM Absence_Breakdown order by 1
SELECT Duration, * FROM Absence

--select dbo.udfcustom_AbsenceDurationForAppointment(GETDATE(), 'AM', getdate(), 'PM', 828)		
--select dbo.udfcustom_AbsenceDurationForAppointment(GETDATE()-1, 'AM', getdate()+1, 'PM', 954)		
