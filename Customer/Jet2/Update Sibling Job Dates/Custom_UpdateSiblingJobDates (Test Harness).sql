DELETE FROM Salary where ID_1 = 2703
INSERT Salary(id_1, Salary_Start_Date, Salary_End_Date, Job_Start_Date,Job_End_Date, Amount) values (2703, '2011-01-01', '2011-03-31', '2011-01-01', '2011-12-31', 10000)
INSERT Salary(id_1, Salary_Start_Date, Salary_End_Date, Job_Start_Date,Job_End_Date, Amount) values (2703, '2012-01-01', '2012-01-31', '2012-01-01', '2012-12-31', 20000)
INSERT Salary(id_1, Salary_Start_Date, Salary_End_Date, Job_Start_Date,Job_End_Date, Amount) values (2703, '2013-01-01', '2013-01-31', '2013-01-01', '2013-12-31', 30000)
INSERT Salary(id_1, Salary_Start_Date, Salary_End_Date, Job_Start_Date,Job_End_Date, Amount) values (2703, '2014-01-01', NULL, '2014-01-01', NULL, 50000)
SELECT id, Job_Start_Date,Job_End_Date, Amount, Notes FROM Salary where ID_1 = 2703 order by amount desc


INSERT Salary(id_1, Salary_Start_Date, Salary_End_Date, Job_Start_Date,Job_End_Date, Amount) values (2703, '2013-04-01', '2014-03-31', '2013-06-10', '2013-06-25', 35000)
SELECT  id, Job_Start_Date,Job_End_Date, Amount, Notes FROM Salary where ID_1 = 2703 order by amount desc

DECLARE @maxID integer
 SELECT @maxID = MAX(id) FROM Salary

--select  id, Job_Start_Date,Job_End_Date, Amount, Notes from Salary WHERE ID = @maxID
UPDATE Salary SET Job_Start_Date = Job_Start_Date - 4 WHERE ID = @maxID

SELECT  id, Job_Start_Date,Job_End_Date, Amount, Notes FROM Salary where ID_1 = 2703 order by amount desc

DELETE Salary WHERE ID = @maxID

--DELETE Salary WHERE ID = 44580
SELECT  id, Job_Start_Date,Job_End_Date, Amount, Notes FROM Salary where ID_1 = 2703 order by amount desc
