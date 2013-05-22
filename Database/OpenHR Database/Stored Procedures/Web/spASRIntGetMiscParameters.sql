CREATE PROCEDURE [dbo].[spASRIntGetMiscParameters]
	(
		@psParam1		varchar(MAX)	OUTPUT,
		@psParam2		varchar(MAX)	OUTPUT,
		@psParam3		varchar(MAX)	OUTPUT,
		@psParam4		varchar(MAX)	OUTPUT
	)
AS
BEGIN
	
	SET @psParam1 = 1;
	SET @psParam2 = 3;
	SET @psParam3 = 300;
	SET @psParam4 = 3600;
	
	SELECT @psParam1 = [ASRSysSystemSettings].[SettingValue]
	FROM [dbo].[ASRSysSystemSettings]
	WHERE [Section] = 'misc' AND [SettingKey] = 'cfg_pcl';
          
	SELECT @psParam2 = [ASRSysSystemSettings].[SettingValue]
	FROM [dbo].[ASRSysSystemSettings]
	WHERE [Section] = 'misc' AND [SettingKey] = 'cfg_ba';

	SELECT @psParam3 = [ASRSysSystemSettings].[SettingValue]
	FROM [dbo].[ASRSysSystemSettings]
	WHERE [Section] = 'misc' AND [SettingKey] = 'cfg_ld';
				
	SELECT @psParam4 = [SettingValue]
	FROM [dbo].[ASRSysSystemSettings]
	WHERE [Section] = 'misc' AND [SettingKey] = 'cfg_rt';

END