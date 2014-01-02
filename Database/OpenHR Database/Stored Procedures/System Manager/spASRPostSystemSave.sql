CREATE PROCEDURE [dbo].[spASRPostSystemSave]
AS
BEGIN

	IF OBJECT_ID('ASRSysProtectsCache') IS NOT NULL 
		DELETE FROM ASRSysProtectsCache;

	INSERT ASRSysProtectsCache 
	SELECT ID, Action, Columns, ProtectType , uid
       FROM sysprotects;

END
