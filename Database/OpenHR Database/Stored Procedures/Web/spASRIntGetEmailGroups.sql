CREATE PROCEDURE [dbo].[spASRIntGetEmailGroups]
AS
BEGIN

	SET NOCOUNT ON;

	SELECT emailGroupID, 
		name, 
		userName, 
		access 
	FROM ASRSysEmailGroupName 
	ORDER BY [name];
END