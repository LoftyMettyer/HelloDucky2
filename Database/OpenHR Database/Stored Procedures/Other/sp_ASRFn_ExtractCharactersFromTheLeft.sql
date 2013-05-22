CREATE PROCEDURE [dbo].[sp_ASRFn_ExtractCharactersFromTheLeft]
(
	@psResult 				varchar(MAX) OUTPUT,
	@psWholeString 			varchar(MAX),
	@piNumberOfCharacters	integer
)
AS
BEGIN
	SET @psResult = LEFT(@psWholeString, @piNumberOfCharacters);
END
