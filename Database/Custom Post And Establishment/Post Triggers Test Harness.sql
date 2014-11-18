
--select * from Post_Records
DECLARE @newID integer;

INSERT Post_Records (Effective_Date) VALUES (GETDATE())

SELECT TOP 1 @newID = ID FROM Post_Records ORDER BY ID DESC
 
SELECT * FROM Post_Records WHERE ID = @newID
SELECT * FROM Post_Holiday_Schemes WHERE ID_219 = @newID
SELECT * FROM Post_OMP_Schemes WHERE ID_219 = @newID
SELECT * FROM Post_OSP_Schemes WHERE ID_219 = @newID
SELECT * FROM Post_Pension_Schemes WHERE ID_219 = @newID
SELECT * FROM Post_Working_Patterns WHERE ID_219 = @newID



