CREATE PROCEDURE [dbo].[spASRIntGetEmailGroups]
AS
BEGIN
	SELECT emailGroupID, 
		name, 
		userName, 
		access 
	FROM ASRSysEmailGroupName 
	ORDER BY [name];
END