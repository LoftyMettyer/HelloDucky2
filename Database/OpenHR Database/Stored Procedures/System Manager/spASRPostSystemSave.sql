CREATE PROCEDURE [dbo].[spASRPostSystemSave]
AS
BEGIN

   SET NOCOUNT ON;

	IF OBJECT_ID('ASRSysProtectsCache') IS NOT NULL 
		DELETE FROM ASRSysProtectsCache;

	INSERT ASRSysProtectsCache ([ID], [Action], [Columns], [ProtectType], [UID])
		SELECT p.ID, Action, Columns, ProtectType , p.uid
			FROM sys.sysprotects p
			INNER JOIN sys.sysobjects o ON o.id = p.id
			WHERE o.xtype = 'V'
			ORDER BY p.uid, name;

END
