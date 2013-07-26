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

DECLARE @lockSystemObjects bit = 1,
				@effectivedate datetime = DATEADD(dd, 0, DATEDIFF(dd, 0, GETDATE()));

MERGE INTO ASRSysColours AS Target
USING (VALUES
	('1', '16777215', 'White', '8', '0'),
	('2', '16777164', 'Light Turquoise', '3', '1'),
	('3', '13434828', 'Light Green', '4', '1'),
	('4', '13434879', 'Light Yellow', '7', '1'),
	('5', '16764057', 'Pale Blue', '3', '1'),
	('6', '16751052', 'Lavender', '5', '1'),
	('7', '13408767', 'Rose', '5', '1'),
	('8', '10079487', 'Tan', '7', '1'),
	('9', '12632256', 'Grey 25%', '16', '0'),
	('10', '16776960', 'Turquoise', '3', '1'),
	('11', '16711935', 'Pink', '5', '1'),
	('12', '65535', 'Yellow', '7', '1'),
	('13', '16763904', 'Sky Blue', '3', '1'),
	('14', '13421619', 'Aqua', '3', '1'),
	('15', '52479', 'Gold', '7', '1'),
	('16', '9868950', 'Grey 40%', '15', '0'),
	('17', '16737843', 'Light Blue', '2', '1'),
	('18', '39423', 'Light Orange', '6', '1'),
	('19', '8421504', 'Grey 50%', '15', '0'),
	('20', '13395456', 'Blue Grey', '2', '1'),
	('21', '52377', 'Lime', '11', '1'),
	('22', '26367', 'Orange', '6', '1'),
	('23', '6723891', 'Sea Green', '11', '1'),
	('24', '6697881', 'Plum', '12', '1'),
	('25', '16711680', 'Blue', '2', '1'),
	('26', '8421376', 'Teal', '10', '1'),
	('27', '8388736', 'Violet', '12', '1'),
	('28', '10040115', 'Indigo', '12', '1'),
	('29', '32896', 'Dark Yellow', '14', '1'),
	('30', '65280', 'Bright Green', '4', '1'),
	('31', '255', 'Red', '6', '1'),
	('32', '13209', 'Brown', '1', '1'),
	('33', '6697728', 'Dark Teal', '1', '1'),
	('34', '8388608', 'Dark Blue', '9', '1'),
	('35', '32768', 'Green', '11', '1'),
	('36', '128', 'Dark Red', '13', '1'),
	('37', '0', 'Black', '1', '0'),
	('38', '6697779', 'Midnight Blue', '3', '0'),
	('39', '16248553', 'Dolphin Blue', '3', '0'),
	('41', '15988214', 'Pale Grey', '15', '0'))
AS Source (ColOrder, ColValue, ColDesc, WordColourIndex, CalendarLegendColour)
ON Target.ColOrder = Source.ColOrder
WHEN MATCHED THEN
	UPDATE SET ColValue = source.colValue, ColDesc = source.ColDesc, WordColourIndex = source.WordColourIndex, CalendarLegendColour = source.CalendarLegendColour
WHEN NOT MATCHED BY TARGET THEN
	INSERT (ColOrder, ColValue, ColDesc, WordColourIndex, CalendarLegendColour)
	VALUES (ColOrder, ColValue, ColDesc, WordColourIndex, CalendarLegendColour)
WHEN NOT MATCHED BY SOURCE THEN 
	DELETE;

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

MERGE INTO tbstat_effectivedates AS Target
USING (VALUES
	(1, @effectivedate)
)
AS Source (type, date)
ON target.type = source.type
WHEN NOT MATCHED BY TARGET THEN
	INSERT (type, date)
	VALUES (type, date)
WHEN NOT MATCHED BY SOURCE THEN 
DELETE;



/* --------------------------------------------------------------------------------------
Customisation data
-------------------------------------------------------------------------------------- */

MERGE INTO tbsys_scriptedobjects AS Target
USING (VALUES 
	('3F55669B-FE5C-4CBA-8E3B-741109FCCD56', 1, 1, 'BB716F9C-B559-4B72-99BE-C5737FC6EE8A', 1, @lockSystemObjects),
	('36FEAB19-D98C-436E-A0BA-23F8E122D709', 1, 2, 'BB716F9C-B559-4B72-99BE-C5737FC6EE8A', 1, @lockSystemObjects),
	('CFF78AA3-1986-4E47-8F25-7086AF99A6BE', 1, 3, 'BB716F9C-B559-4B72-99BE-C5737FC6EE8A', 1, @lockSystemObjects)
)
AS Source (guid, objecttype, targetid, ownerid, revision, locked)
ON Target.guid = Source.guid
WHEN MATCHED THEN 
	UPDATE SET objecttype = source.objecttype, targetid = source.targetid, ownerid = source.ownerid, revision = source.revision
WHEN NOT MATCHED BY TARGET THEN 
	INSERT (guid, objecttype, targetid, ownerid, revision, locked, effectivedate)
	VALUES (guid, objecttype, targetid, ownerid, revision, locked, @effectivedate)
WHEN NOT MATCHED BY SOURCE THEN 
DELETE;

MERGE INTO tbsys_tables AS Target 
USING (VALUES 
	('1', '1', 'Personnel_Records'), 
	('2', '2', 'Absence'), 
	('3', '2', 'Salary')
) 
AS Source (tableID, tableType, tableName) 
ON Target.tableID = Source.tableID AND Target.tablename = Source.tablename
WHEN MATCHED THEN 
	UPDATE SET tabletype = Source.tabletype
WHEN NOT MATCHED BY TARGET THEN 
	INSERT (tableID, tableType, tableName, DefaultOrderID, RecordDescExprID, DefaultEmailID, ManualSummaryColumnBreaks, AuditInsert, AuditDelete, IsRemoteView) 
	VALUES (tableID, tableType, tableName, 0, 0, 0, 0, 0, 0, 0) 
WHEN NOT MATCHED BY SOURCE THEN 
	DELETE;