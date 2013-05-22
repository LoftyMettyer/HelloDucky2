CREATE PROCEDURE [dbo].[sp_ASRFn_IsPopulated_1]
(
	@pfResult	bit OUTPUT,
	@psString	varchar(MAX)
)
AS
BEGIN
	SET @pfResult = 1;

	IF LEN(@psString) = 0 SET @pfResult = 0;

	IF @psString IS NULL SET @pfResult = 0;

END
