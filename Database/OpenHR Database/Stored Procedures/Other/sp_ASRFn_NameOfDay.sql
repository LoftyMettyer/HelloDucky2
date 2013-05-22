CREATE PROCEDURE [dbo].[sp_ASRFn_NameOfDay]
(
	@psResult	varchar(MAX) OUTPUT,
	@pdtDate 	datetime
)
AS
BEGIN
	SET @psResult = DATENAME(weekday, @pdtDate);
END
