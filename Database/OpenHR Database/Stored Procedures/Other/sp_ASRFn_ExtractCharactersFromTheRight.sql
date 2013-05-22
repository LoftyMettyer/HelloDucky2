CREATE PROCEDURE [dbo].[sp_ASRFn_ExtractCharactersFromTheRight]
(
	@psResult 				varchar(MAX) OUTPUT,
	@psWholeString 			varchar(MAX),
	@piNumberOfCharacters	integer
)
AS
BEGIN
	SET @psResult = RIGHT(@psWholeString, @piNumberOfCharacters);
END
