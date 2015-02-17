CREATE PROCEDURE spASRIntSaveExpression (
	@expressionID integer = 0 OUTPUT,
	@name varchar(255),
	@TableID integer,
	@returnType integer,
	@returnSize integer,
	@returnDecimals integer,
	@type integer,
	@parentComponentID integer,
	@Username varchar(50) = '',
	@access varchar(2) = '',
	@description varchar(255) = '')
AS
BEGIN

	SET NOCOUNT ON;

	IF @expressionID = 0
	BEGIN

		-- Keep a manual record of allocated IDs in case users in SYS MGR have created expressions but not yet saved changes
		SELECT @expressionID = ISNULL([SettingValue], 1) FROM [asrsyssystemsettings] WHERE [Section] = 'AUTOID' AND [SettingKey] = 'expressions';
		IF @expressionID = 1
			INSERT ASRSysSystemSettings([Section], [SettingKey], [SettingValue]) VALUES ('AUTOID', 'ExprID', 1);	
		ELSE
		BEGIN
			SET @expressionID = @expressionID + 1;
			UPDATE ASRSysSystemSettings SET [SettingValue] = @expressionID WHERE [Section] ='AUTOID' AND [SettingKey] = 'expressions';
		END

		INSERT ASRSysExpressions (exprID, name, TableID, returnType, returnSize, returnDecimals, [type], parentComponentID, Username, access, [description])
			VALUES(@expressionID, @name, @TableID, @returnType, @returnSize, @returnDecimals, @type, @parentComponentID, @Username, @access, @description);
	END
	ELSE
	BEGIN
		UPDATE ASRSysExpressions SET name = @name, TableID = @TableID, returnType = @returnType, returnSize = @returnSize, returnDecimals = @returnDecimals,
				[type] = @type, parentComponentID = @parentComponentID, Username = @Username, access = @access, [description] = @description
					WHERE ExprID = @expressionID;
	END

END
