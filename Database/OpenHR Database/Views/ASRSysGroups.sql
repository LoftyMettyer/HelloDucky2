CREATE VIEW [dbo].[ASRSysGroups] AS
	SELECT principal_id AS ID, name FROM sys.database_principals
	WHERE type = 'R' AND is_fixed_role = 0
		AND (principal_id > 0) AND NOT (name LIKE 'ASRSys%');
