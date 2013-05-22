CREATE PROCEDURE [dbo].[sp_ASROp_Or]
(
	@pfResult	bit OUTPUT,
	@pfFirst	bit,
	@pfSecond	bit
)
AS
BEGIN
	SET @pfResult = @pfFirst | @pfSecond;
END