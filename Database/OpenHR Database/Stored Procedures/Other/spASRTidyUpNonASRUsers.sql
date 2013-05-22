CREATE PROCEDURE [dbo].[spASRTidyUpNonASRUsers]
AS
BEGIN
	SET NOCOUNT ON;
	
	DECLARE @iGroupID int;

	SELECT @iGroupID = uid FROM sysusers WHERE isSQLRole = 1 AND name = 'ASRSysGroup';

	SELECT * FROM sysusers WHERE uid NOT IN (SELECT uid FROM sysmembers
	INNER JOIN sysusers ON sysusers.uid = sysmembers.memberuid
	WHERE groupuid = @iGroupID)
	AND IsSQLRole = 0 AND NOT (name = 'dbo' OR name= 'guest' OR name='sys' OR name='INFORMATION_SCHEMA');

	RETURN;

END