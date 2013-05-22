CREATE PROCEDURE [dbo].[sp_ASRFn_IsEmpty_1]
(
	@pfResult	bit OUTPUT,
	@psString	varchar(MAX)
)
AS
BEGIN
	SET @pfResult = 0;

	IF LEN(@psString) = 0 
		SET @pfResult = 1;

	IF @psString IS null
		SET @pfResult = 1;

END
