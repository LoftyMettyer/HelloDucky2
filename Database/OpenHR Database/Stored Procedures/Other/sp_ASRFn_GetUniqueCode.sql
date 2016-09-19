CREATE PROCEDURE [dbo].[sp_ASRFn_GetUniqueCode]
(
   @piInstanceID	int,
	@piResult		int OUTPUT,
	@psCodePrefix	varchar(MAX) = '',
	@piSuffixRoot	int=1
)
AS
BEGIN
	SELECT @piResult = [dbo].[udfstat_getuniquecode] (@psCodePrefix, @piSuffixRoot, @piInstanceID);
END