CREATE PROCEDURE [dbo].[sp_ASRFn_ExtractPartOfAString]
(
	@psResult 				varchar(MAX) OUTPUT,
	@psString 				varchar(MAX),
	@piStart 				integer,
	@piNumberOfCharacters	integer
)
AS
BEGIN
	SET @psResult = SUBSTRING(@psString, @piStart, @piNumberOfCharacters);
END
