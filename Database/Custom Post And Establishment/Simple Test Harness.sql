--SELECT id,* FROM Appointments WHERE ID = 825
--SELECT id,* FROM Appointment_Working_Patterns

DELETE FROM Absence_Entry
DELETE FROM Appointment_Absence_Entry
DELETE FROM Absence_Breakdown
DELETE FROM Absence

    /* This sets all of the flags prior to updating date dependant columns */
    DELETE FROM ASRSYSSystemSettings WHERE [Section] = 'database' and [SettingKey] = 'updatingdatedependantcolumns'

    INSERT ASRSYSSystemSettings([Section],[SettingKey],[SettingValue])
    VALUES('database','updatingdatedependantcolumns',1)

print 'iNSERT'
INSERT Appointment_Absence_Entry (ID_3, start_date, start_session, end_date, end_session, Reason, absence_type) values (825, getdate()-1000, 'AM', getdate()-995, 'PM', 'multiple post day off', 'HOLS')
INSERT Appointment_Absence_Entry (ID_3, start_date, start_session, end_date, end_session, Reason, absence_type) values (825, getdate()-1, 'AM', getdate()-1, 'PM', 'just a single day off', 'MAT')
INSERT Appointment_Absence_Entry (ID_3, start_date, start_session, end_date, end_session, Reason, absence_type) values (825, getdate(), 'AM', getdate(), 'AM', 'morning', 'HOLS')
INSERT Appointment_Absence_Entry (ID_3, start_date, start_session, end_date, end_session, Reason, absence_type) values (825, getdate()+1, 'PM', getdate()+1, 'PM', 'afternoon', 'JURY')
INSERT Appointment_Absence_Entry (ID_3, start_date, start_session, end_date, end_session, Reason, absence_type) values (825, getdate()-2, 'PM', getdate()+7, 'AM', 'afternoon', 'JURY')
INSERT Absence_Entry (ID_1, start_date, end_date, Reason, absence_type) values (105, getdate(), getdate()+1, 'SICK', 'felt bad')

PRINT 'UPDATE'
--UPDATE Absence_Entry SET start_date = start_date - 1
--UPDATE Appointment_Absence_Entry SET start_date = end_date - 1
-- DELETE FROM Absence_Entry --WHERE ID = 182
-- DELETE FROM Appointment_Absence_Entry --WHERE ID = 86
--SELECT * FROM Absence_Entry
select * from Appointment_Working_Patterns

SELECT * FROM Appointment_Absence_Entry
SELECT Session,* FROM Absence_Breakdown
SELECT id, Start_date, End_Date, * FROM Absence
