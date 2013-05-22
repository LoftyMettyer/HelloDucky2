CREATE PROCEDURE sp_ASRFn_SystemDate 
(
	@pdtResult	datetime OUTPUT
)
AS
BEGIN
	/* Get the current system date. */
	SET @pdtResult = convert(datetime, convert(varchar(20), getdate(), 101))
END
GO

