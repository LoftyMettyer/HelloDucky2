CREATE PROCEDURE [dbo].[sp_ASRUniqueObjectName](
		  @psUniqueObjectName sysname OUTPUT
		, @Prefix sysname
		, @Type int)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE @NewObj 		as sysname
		, @Count 			as integer
		, @sUserName		as sysname
		, @sCommandString	nvarchar(MAX)	
 		, @sParamDefinition	nvarchar(500);

	SET @sUserName = SYSTEM_USER;
	SET @Count = 1;
	SET @NewObj = @Prefix + CONVERT(varchar(100),@Count);

	WHILE (EXISTS (SELECT * FROM sysobjects WHERE id = object_id(@NewObj) AND sysstat & 0xf = @Type))
		OR (EXISTS (SELECT * FROM ASRSysSQLObjects WHERE Name = @NewObj AND Type = @Type))
		BEGIN
			SET @Count = @Count + 1;
			SET @NewObj = @Prefix + CONVERT(varchar(10),@Count);
		END

	INSERT INTO [dbo].[ASRSysSQLObjects] ([Name], [Type], [DateCreated], [Owner])
		VALUES (@NewObj, @Type, GETDATE(), @sUserName);

	SET @sCommandString = 'SELECT @psUniqueObjectName = ''' + @NewObj + '''';
	SET @sParamDefinition = N'@psUniqueObjectName sysname output';
	EXECUTE sp_executesql @sCommandString, @sParamDefinition, @psUniqueObjectName output;

END