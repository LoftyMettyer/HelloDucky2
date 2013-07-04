/*
Post-Deployment Script Template							
--------------------------------------------------------------------------------------
 This file contains SQL statements that will be appended to the build script.		
 Use SQLCMD syntax to include a file in the post-deployment script.			
 Example:      :r .\myfile.sql								
 Use SQLCMD syntax to reference a variable in the post-deployment script.		
 Example:      :setvar TableName MyTable							
               SELECT * FROM [$(TableName)]					
--------------------------------------------------------------------------------------
*/

-- Reference Data for AddressType 
MERGE INTO ASRSysSystemSettings AS Target 
USING (VALUES 
  ('database', N'olestructure','2'), 
  ('database', N'ownerid','BB716F9C-B559-4B72-99BE-C5737FC6EE8A'), 
  ('database', N'version','5.2')
) 
AS Source (Section, SettingKey, SettingValue) 
ON Target.Section = Source.Section AND Target.SettingKey = Source.SettingKey
-- update matched rows 
WHEN MATCHED THEN 
UPDATE SET SettingValue = Source.SettingValue 
-- insert new rows 
WHEN NOT MATCHED BY TARGET THEN 
INSERT (Section, SettingKey, SettingValue) 
VALUES (Section, SettingKey, SettingValue) 
-- delete rows that are in the target but not the source 
WHEN NOT MATCHED BY SOURCE THEN 
DELETE;
