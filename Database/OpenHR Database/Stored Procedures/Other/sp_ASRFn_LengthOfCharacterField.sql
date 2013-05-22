CREATE PROCEDURE [dbo].[sp_ASRFn_LengthOfCharacterField]
(
	@piResult 		integer OUTPUT,
	@psWholeString	varchar(MAX)
)
AS
BEGIN
	SET @piResult = LEN(@psWholeString);
END
