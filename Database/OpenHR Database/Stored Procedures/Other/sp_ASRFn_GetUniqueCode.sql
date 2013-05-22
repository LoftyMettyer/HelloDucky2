CREATE PROCEDURE [dbo].[sp_ASRFn_GetUniqueCode]
(
	@piResult		int OUTPUT,
	@psCodePrefix	varchar(MAX) = '',
	@piSuffixRoot	int=1
)
AS
BEGIN
	DECLARE @iOldCodeSuffix int;
	DECLARE @iNewCodeSuffix int;

	-- Get the current maximum code suffix for the given code prefix.
	SELECT @iOldCodeSuffix = maxCodeSuffix 
		FROM [dbo].[tbsys_uniquecodes]
		WHERE codePrefix = @psCodePrefix;

	IF @iOldCodeSuffix IS NULL 
	BEGIN
		-- The given code prefix DOES NOT exist in the database, so set the suffix to be the given root suffix, and insert the new code into the database.
		SELECT @iNewCodeSuffix = @piSuffixRoot;
		INSERT INTO [dbo].[tbsys_uniquecodes] (codePrefix, maxCodeSuffix) VALUES (@psCodePrefix, @iNewCodeSuffix);
	END
	ELSE
	BEGIN
		-- The given code prefix DOES exist in the database, so set the suffix to be the current max suffix plus 1, and update the code into the database.
		SELECT @iNewCodeSuffix = @iOldCodeSuffix + 1;
		UPDATE [dbo].[tbsys_uniquecodes] SET maxCodeSuffix = @iNewCodeSuffix WHERE codePrefix = @psCodePrefix;
	END

	-- Return the new code suffix.
	SET @piResult = @iNewCodeSuffix;
END
