DECLARE @newID integer;

INSERT Appointments (Appointment_Start_Date, Appointment_End_Date, ID_219, id_1) VALUES (GETDATE()-365, NULL, 221, 139)
--update Appointments set actual_hours = actual_hours
--INSERT Post_Allowances (ID_219, Type, Frequency, Amount, Currency) VALUES (164, 'mytype', 3, 333, 'c')
--select * from Post_Allowances
--select * from Post_Records

SELECT TOP 1 @newID = ID FROM Appointments ORDER BY ID DESC
 
SELECT * FROM Appointments WHERE ID = @newID

SELECT * FROM Appointment_Allowances WHERE ID_3 = @newID
SELECT * FROM Appointment_Benefits WHERE ID_3 = @newID
SELECT * FROM Appointment_Deductions WHERE ID_3 = @newID
SELECT * FROM Appointment_Holiday_Schemes WHERE ID_3 = @newID
SELECT * FROM Appointment_OMP_Schemes WHERE ID_3 = @newID
SELECT * FROM Appointment_OSP_Schemes WHERE ID_3 = @newID
SELECT * FROM Appointment_Pension_Schemes WHERE ID_3 = @newID
SELECT * FROM Appointment_Working_Patterns WHERE ID_3 = @newID

-- SELECT * FROM Post_Records where id = 40
--SELECT * FROM Post_Pension_Schemes
--select * from Post_Deductions
select * from Working_Patterns order by id_1

