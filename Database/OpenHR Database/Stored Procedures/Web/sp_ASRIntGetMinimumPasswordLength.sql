CREATE PROCEDURE [dbo].[sp_ASRIntGetMinimumPasswordLength] (
	@piMinPassordLength	integer		OUTPUT
)
AS
BEGIN
	/* Return the minimum password length. */
	DECLARE 
		@sValue				varchar(MAX),
		@fNewSettingFound	bit,
		@fOldSettingFound	bit;

	/* Get the minimum password length. */
	SET @piMinPassordLength = 0;
	exec sp_ASRIntGetSystemSetting 'password', 'minimum length', 'minimumPasswordLength', @sValue OUTPUT, @fNewSettingFound OUTPUT, @fOldSettingFound OUTPUT;
	IF (@fNewSettingFound = 1) OR (@fOldSettingFound = 1) 
	BEGIN
		SET @piMinPassordLength = convert(integer, @sValue);
	END

	IF @piMinPassordLength IS NULL SET @piMinPassordLength = 0;
END