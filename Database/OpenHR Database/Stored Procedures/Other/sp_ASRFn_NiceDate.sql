CREATE PROCEDURE [dbo].[sp_ASRFn_NiceDate]
(
	@psResult	varchar(MAX) OUTPUT,
	@pdtDate 	datetime
)
AS
BEGIN
	
	-- Format(pvParam1, "dddd, mmmm d yyyy")
	IF @pdtDate IS NULL
	BEGIN
		SET @psResult = '';
	END
	ELSE
	BEGIN
		SET @psResult = datename(dw, @pdtDate) + ', ' + datename(mm, @pdtDate) + ' ' + ltrim(str(datepart(dd, @pdtDate))) + ' ' + ltrim(str(datepart(yy, @pdtDate)));
	END
END
