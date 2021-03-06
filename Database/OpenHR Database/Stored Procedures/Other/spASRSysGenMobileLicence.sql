CREATE PROCEDURE [dbo].[spASRSysGenMobileLicence](
  @piLicenceQty bigint
  ) 
AS
BEGIN
  DECLARE 
  @sGUID varchar(MAX),
  @iCount integer;

	SELECT @iCount = COUNT(*) FROM ASRSysSystemSettings WHERE [Section] = 'licence' AND [SettingKey] = 'mobile';
	
	IF @iCount > 0 DELETE FROM ASRSysSystemSettings WHERE [Section] = 'licence' AND [SettingKey] = 'mobile';

	IF @piLicenceQty <= 0 RETURN;

	SET @sGUID = NEWID();
	
	SET @sGUID = @sGUID + '-EA' + CONVERT(VARCHAR(MAX), @piLicenceQty) + 'FF';
	
	INSERT INTO ASRSysSystemSettings
		(Section, SettingKey, SettingValue)
		VALUES 
		('licence', 'mobile', @sGUID);

END;