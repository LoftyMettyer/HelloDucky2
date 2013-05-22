CREATE PROCEDURE [dbo].[spASRIntGetEventLogEmails]
AS
BEGIN

  SELECT [ASRSysEmailGroupName].[EmailGroupID] AS 'EmailGroupID', 
				 [ASRSysEmailGroupName].[Name] AS 'Name'
  FROM [ASRSysEmailGroupName]
  UNION
  SELECT -1  AS 'EmailGroupID',
				(SELECT [ASRSysSystemSettings].[SettingValue]
          FROM [ASRSysSystemSettings]
          WHERE ([ASRSysSystemSettings].[Section] = 'Support')
             AND ([ASRSysSystemSettings].[SettingKey] = 'Email')
         ) AS 'Name'
  ORDER BY 'Name';

END