--select id, * from post_records
--select * from post_working_patterns
DECLARE @persID integer

SELECT @persID = ID FROM Personnel_Records WHERE surname = 'Bourne'
--select @persID

--SELECT * FROM Post_Records where ID IN (58, 65)
--SELECT * FROM Post_Working_Patterns where ID_219 IN (58, 65)
--DELETE FROM Post_Working_Patterns where ID > 300




--SELECT * FROM Post_Working_Patterns
--select max(id) from appointments where id_219 =
--select * from Appointment_Working_Patterns where id_3 >= 898
-- select * from appointments where id_3 = 114
DELETE FROM Appointments WHERE ID_1 = @persID
--delete from Appointment_Working_Patterns
delete from Working_Patterns WHERE ID_1 = @persID
--delete from appointment_Working_Patterns WHERE ID_3 in (58, 65)
 --delete from post_working_patterns where id >= 283

--SELECT TOP 1 @newID = ID FROM Appointment_Working_Patterns ORDER BY ID DESC

DECLARE @newID integer;

-- Insert appointment #1
INSERT Appointments (Appointment_Start_Date, ID_219, ID_1) VALUES (GETDATE(), 65, @persID)

-- Update default working pattern
SELECT TOP 1 @newID = ID FROM Appointment_Working_Patterns ORDER BY ID DESC
UPDATE Appointment_Working_Patterns SET effective_date = effective_date + 10 where id = @newID

-- Insert a new working pattern (future)
SELECT TOP 1 @newID = ID FROM Appointments ORDER BY ID DESC
INSERT Appointment_Working_Patterns (Effective_Date, End_Date, Absence_In, ID_3, Monday_Hours_AM, Tuesday_Hours_AM, Wednesday_Hours_AM, Thursday_Hours_AM) VALUES ('2016-01-01', '2016-12-31', 'Days', @newID, 1, 1, 1, 1)



-- Insert Appointment #2
INSERT Appointments (Appointment_Start_Date, ID_219, ID_1) VALUES (GETDATE(), 58, @persID)
SELECT TOP 1 @newID = ID FROM Appointments ORDER BY ID DESC


--DELETE FROM Appointment_Working_Patterns WHERE ID_3 = @newID

--INSERT Appointment_Working_Patterns (Effective_Date, End_Date, Absence_In, ID_3, Saturday_Hours_AM) VALUES ('2015-01-01', NULL, 'Hours', @newID, 4.33)
--update Post_Records set Post_ID=' ducky162' where ID = 162
--select * from Post_Records where ID = 164

--INSERT Appointments (Appointment_Start_Date, ID_219, ID_1, Post_ID) VALUES (GETDATE(), 164, 117, 'ducky 2')
--SELECT TOP 1 @newID = ID FROM Appointments ORDER BY ID DESC
--DELETE FROM Appointment_Working_Patterns WHERE ID_3 = @newID
--INSERT Appointment_Working_Patterns (Effective_Date, End_Date, ID_3, Tuesday_Hours_AM, Friday_Hours_AM) VALUES ('2015-01-01', '2015-01-31',  @newID, 1, 1)



-- View all appointment working patterns
SELECT awp.* FROM Appointments a
	inner join appointment_working_patterns awp on awp.ID_3 = a.ID
	where a.ID_1 = @persID

-- View merged working patterns
select * from Working_Patterns 
	where id_1 = @persID
	order by Effective_Date DESC
