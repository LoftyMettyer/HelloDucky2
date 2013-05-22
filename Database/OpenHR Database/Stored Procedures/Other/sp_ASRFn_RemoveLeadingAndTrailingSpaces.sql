CREATE PROCEDURE [dbo].[sp_ASRFn_RemoveLeadingAndTrailingSpaces]
(
	@psResult		varchar(MAX) OUTPUT,
	@psTextToTrim	varchar(MAX)
)
AS
BEGIN
	SET @psResult = LTRIM(RTRIM(@psTextToTrim));
END