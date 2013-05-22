CREATE PROCEDURE [dbo].[sp_ASRFn_NameOfMonth]
(
	@psResult	varchar(MAX) OUTPUT,
	@pdtDate 	datetime
)
AS
BEGIN
	SET @psResult = DATENAME(month, @pdtDate);
END
