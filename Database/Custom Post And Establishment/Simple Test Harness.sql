--select id_1 from Appointments order by 1
--select id from Personnel_Records where surname = 'Foster'

--SELECT id,* FROM Appointments WHERE ID_1 = 116
--
SELECT id,* FROM Appointment_Working_Patterns WHERE ID_3 = 876
--delete from Appointment_Working_Patterns where ID in (130)
--UPDATE Appointment_Working_Patterns SET Effective_Date = '2014-11-17' WHERE ID = 126
--UPDATE Appointment_Working_Patterns SET friDAY_HOURS_aM = 1, fridaY_hOURS_pM = 2  WHERE ID = 126

DELETE FROM Absence_Entry
DELETE FROM Appointment_Absence_Entry
DELETE FROM Absence_Breakdown
DELETE FROM Absence

    /* This sets all of the flags prior to updating date dependant columns */
    DELETE FROM ASRSYSSystemSettings WHERE [Section] = 'database' and [SettingKey] = 'updatingdatedependantcolumns'

    INSERT ASRSYSSystemSettings([Section],[SettingKey],[SettingValue])
    VALUES('database','updatingdatedependantcolumns',1)

print 'iNSERT'
--INSERT Appointment_Absence_Entry (ID_3, start_date, start_session, end_date, end_session, Reason, absence_type) values (825, getdate()-1000, 'AM', getdate()-995, 'PM', 'multiple post day off', 'HOLS')
--INSERT Appointment_Absence_Entry (ID_3, start_date, start_session, end_date, end_session, Reason, absence_type) values (825, getdate()-1, 'AM', getdate()-1, 'PM', 'just a single day off', 'MAT')
--INSERT Appointment_Absence_Entry (ID_3, start_date, start_session, end_date, end_session, Reason, absence_type) values (825, getdate(), 'AM', getdate(), 'AM', 'morning', 'HOLS')
--INSERT Appointment_Absence_Entry (ID_3, start_date, start_session, end_date, end_session, Reason, absence_type) values (825, getdate()+1, 'PM', getdate()+1, 'PM', 'afternoon', 'JURY')
INSERT Appointment_Absence_Entry (ID_3, start_date, start_session, end_date, end_session, Reason, absence_type) values (876, getdate()-5, 'PM', getdate()-5, 'AM', 'afternoon', 'JURY')
--INSERT Absence_Entry (ID_1, start_date, end_date, Reason, absence_type) values (105, getdate(), getdate(), 'SICK', 'felt bad')

INSERT Absence_Entry (ID_1, start_date, start_session, end_date, end_session, absence_type, Reason, Absence_In) values (116, getdate()-5, 'PM', getdate()-5, 'AM', 'SICK', 'Jody felt bad', 'Hours')
--INSERT Absence_Entry (ID_1, start_date, start_session, end_date, end_session, absence_type, Reason) values (116, getdate(), 'AM', getdate(), 'PM', 'SICK', 'Jody felt bad')

PRINT 'UPDATE'
--UPDATE Absence_Entry SET start_date = start_date - 1
--UPDATE Appointment_Absence_Entry SET start_date = end_date - 1
-- DELETE FROM Absence_Entry --WHERE ID = 182
-- DELETE FROM Appointment_Absence_Entry --WHERE ID = 86
--SELECT * FROM Absence_Entry
--SELECT id,* FROM Appointment_Working_Patterns where id_3 in (835, 876)
--select * from Appointment_Working_Patterns
--SELECT * FROM Appointment_Absence_Entry
SELECT * FROM Absence_Breakdown
SELECT id, Start_date, Start_Session, End_Date, End_Session, Reason, * FROM Absence
