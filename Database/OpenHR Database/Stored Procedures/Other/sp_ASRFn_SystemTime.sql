CREATE PROCEDURE [dbo].[sp_ASRFn_SystemTime]
(
	@psResult	varchar(MAX) OUTPUT
)
AS
BEGIN
	SET @psResult = convert(varchar(20), GETDATE(), 8);
END
