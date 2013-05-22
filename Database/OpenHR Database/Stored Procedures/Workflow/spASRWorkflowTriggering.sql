CREATE PROCEDURE [dbo].[spASRWorkflowTriggering]
(
	@pfTrigger bit OUTPUT
)
AS
BEGIN
	DECLARE @sInProgress varchar(MAX);

	SET @pfTrigger = 0;

	SELECT @sInProgress = isnull(settingValue, '0')
	FROM ASRSysSystemSettings
	WHERE section = 'workflow'
		AND settingKey = 'triggering';

	IF @sInProgress = '0'
	BEGIN
		SET @pfTrigger = 1;

		DELETE FROM ASRSysSystemSettings
		WHERE section = 'workflow'
			AND settingKey = 'triggering';

		INSERT INTO ASRSysSystemSettings
			(section, settingKey, settingValue)
		VALUES 
			('workflow', 'triggering', '1');
	END
END