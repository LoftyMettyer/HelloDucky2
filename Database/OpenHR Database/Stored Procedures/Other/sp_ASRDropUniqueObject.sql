CREATE PROCEDURE [dbo].[sp_ASRDropUniqueObject](
	@psUniqueObjectName sysname,
	@piType integer)
AS
BEGIN
	DECLARE 
		@sCommandString				nvarchar(MAX),
		@sCleanUniqueObjectName		sysname;

	/* Clean the input string parameters. */
	SET @sCleanUniqueObjectName = @psUniqueObjectName;
	IF len(@sCleanUniqueObjectName) > 0 SET @sCleanUniqueObjectName = replace(@sCleanUniqueObjectName, '''', '''''');
										
	IF (EXISTS (SELECT * 
							FROM sysobjects 
							WHERE name = @psUniqueObjectName))
	BEGIN
		IF @piType = 3 
		BEGIN
			SET @sCommandString = 'DROP TABLE ' + @sCleanUniqueObjectName;
		END

		IF @piType = 4
		BEGIN
			SET @sCommandString = 'DROP PROCEDURE ' + @sCleanUniqueObjectName;
		END 

		EXECUTE sp_executesql @sCommandString;
  END
	
	DELETE FROM [dbo].[ASRSysSQLObjects]
	WHERE [Name] = @psUniqueObjectName 
		AND [Type] = @piType
		AND [Owner] = SYSTEM_USER;

END