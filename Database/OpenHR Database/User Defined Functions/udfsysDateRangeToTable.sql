CREATE FUNCTION [dbo].[udfsysDateRangeToTable]
(     
      @Increment              char(1),
      @StartDate              datetime,
      @StartSession           char(2),
      @EndDate                datetime,
	  @EndSession			  char(2)
)
RETURNS  
	@SelectedRange	TABLE ([IndividualDate] datetime, [SessionType] char(3))
AS 
BEGIN
	SET @StartDate = DATEADD(dd, 0, DATEDIFF(dd, 0, @StartDate));
	SET @EndDate = DATEADD(dd, 0, DATEDIFF(dd, 0, @EndDate));

    WITH cteRange (DateRange) AS (
        SELECT @StartDate
        UNION ALL
        SELECT DATEADD(dd, 0, DATEDIFF(dd, 0, 
                CASE
                    WHEN @Increment = 'd' THEN DATEADD(dd, 1, DateRange)
                    WHEN @Increment = 'w' THEN DATEADD(ww, 1, DateRange)
                    WHEN @Increment = 'm' THEN DATEADD(mm, 1, DateRange)
                END))
        FROM cteRange
        WHERE DateRange <= 
                CASE
                    WHEN @Increment = 'd' THEN DATEADD(dd, -1, @EndDate)
                    WHEN @Increment = 'w' THEN DATEADD(ww, -1, @EndDate)
                    WHEN @Increment = 'm' THEN DATEADD(mm, -1, @EndDate)
                END)         
    INSERT INTO @SelectedRange (IndividualDate, SessionType)
    SELECT DateRange, 
	CASE
		WHEN @StartSession = 'PM' AND DateRange = @StartDate THEN 'PM'
		WHEN @EndSession = 'AM' AND DateRange = @EndDate THEN 'AM'
		ELSE 'Day'
	END
    FROM cteRange
    OPTION (MAXRECURSION 3660);
    RETURN;
END
GO