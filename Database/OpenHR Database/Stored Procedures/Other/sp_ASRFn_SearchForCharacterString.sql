CREATE PROCEDURE [dbo].[sp_ASRFn_SearchForCharacterString]
(	
	@piResult		integer OUTPUT,
	@psWholeString	varchar(MAX),
	@psSearchString	varchar(MAX)
)
AS
BEGIN
	SET @piResult = CHARINDEX(@psSearchString, @psWholeString);
END
