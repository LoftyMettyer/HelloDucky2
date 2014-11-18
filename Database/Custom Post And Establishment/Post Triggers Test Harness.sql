--INSERT PSG_Holiday_Schemes (Effective_Date, End_Date, Holiday_Scheme, ID_215) VALUES (GETDATE()-100, GETDATE()+100, 'holscheme', 40)
--INSERT PSG_OMP_Schemes (Effective_Date, End_Date, OMP_Scheme, ID_215) VALUES (GETDATE()-100, GETDATE()+100, 'OMP', 40)
--INSERT PSG_OSP_Schemes (Effective_Date, End_Date, OSP_Scheme, ID_215) VALUES (GETDATE()-100, GETDATE()+100, 'osp', 40)
--INSERT PSG_Pension_Schemes (Effective_Date, End_Date, Pension_Scheme, ID_215) VALUES (GETDATE()-100, GETDATE()+100, 'pen1', 40)
--INSERT PSG_Working_Patterns (ID_215, Effective_Date, Regional_ID, Absence_In, Day_Pattern, Sunday_Hours_AM, Sunday_Hours_PM, Monday_Hours_AM, Monday_Hours_PM
--								, Tuesday_Hours_AM, Tuesday_Hours_PM, Wednesday_Hours_AM, Wednesday_Hours_PM, Thursday_Hours_AM, Thursday_Hours_PM
--								, Friday_Hours_AM, Friday_Hours_PM, Saturday_Hours_AM, Saturday_Hours_PM)
--								VALUES (40, GETDATE(), 1, 'Hours', 'SSMMTTWWTTFFSS', 1,2,3,4,5,6,7,8,9,10,11,12,13,14)


--select * from Post_Records

DECLARE @newID integer;

INSERT Post_Records (Effective_Date, Pay_Scale_Group) VALUES (GETDATE(), 'Sales')

SELECT TOP 1 @newID = ID FROM Post_Records ORDER BY ID DESC
 
SELECT * FROM Post_Records WHERE ID = @newID
SELECT * FROM Post_Holiday_Schemes WHERE ID_219 = @newID
SELECT * FROM Post_OMP_Schemes WHERE ID_219 = @newID
SELECT * FROM Post_OSP_Schemes WHERE ID_219 = @newID
SELECT * FROM Post_Pension_Schemes WHERE ID_219 = @newID
SELECT * FROM Post_Working_Patterns WHERE ID_219 = @newID


--SELECT * FROM Pay_Scale_Groups
--SELECT * FROM PSG_Holiday_Schemes
--SELECT * FROM ASRSysTables
--SELECT * FROM PSG_Working_Patterns