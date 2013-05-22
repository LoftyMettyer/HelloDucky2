CREATE FUNCTION udfASRDateOverlap(
	@pdStartDate1		datetime,
	@psStartSession1	nvarchar(2),
	@pdEndDate1			datetime,
	@psEndSession1		nvarchar(2),
	@psType1			nvarchar(MAX),
	@pdStartDate2		datetime,
	@psStartSession2	nvarchar(2),
	@pdEndDate2			datetime,
	@psEndSession2		nvarchar(2),
	@psType2			nvarchar(MAX))
RETURNS bit
AS
BEGIN
	
	DECLARE @bFound bit;
	SET @bFound = 0;
	
	-- 1st of data is the inserted, 2nd is physical database values.	
	SET @pdStartDate1 = DATEADD(D, 0, DATEDIFF(D, 0, @pdStartDate1));
	SET @pdEndDate1 = ISNULL(DATEADD(D, 0, DATEDIFF(D, 0, @pdEndDate1)), CONVERT(datetime,'9999-12-31'));
	SET @pdStartDate2 = DATEADD(D, 0, DATEDIFF(D, 0, @pdStartDate2));
	SET @pdEndDate2 = ISNULL(DATEADD(D, 0, DATEDIFF(D, 0, @pdEndDate2)), CONVERT(datetime,'9999-12-31'));

	-- Put the AM/PM stuff into the above dates.
	IF @psStartSession1 = 'PM' SET @pdStartDate1 = DATEADD(hh, 12, @pdStartDate1);
	IF @psEndSession1 = 'PM' SET @pdEndDate1 = DATEADD(hh, 23, @pdEndDate1);
	IF @psStartSession2 = 'PM' SET @pdStartDate2 = DATEADD(hh, 12, @pdStartDate2);
	IF @psEndSession2 = 'PM' SET @pdEndDate2 = DATEADD(hh, 23, @pdEndDate2);

	-- Check to see if this date overlaps.
	IF ((@pdStartDate1 BETWEEN @pdStartDate2 AND @pdEndDate2)
		OR (@pdEndDate1 BETWEEN @pdStartDate2 AND @pdEndDate2)
		OR (@pdStartDate1 < @pdStartDate2 AND @pdEndDate1 > @pdEndDate2))
		AND (@psType1 = @psType2 OR @psType1 IS NULL)
			SET @bFound = 1;

	RETURN @bFound;

END