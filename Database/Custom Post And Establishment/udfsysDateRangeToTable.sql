CREATE FUNCTION [dbo].[udfsysDateRangeToTable]
(     
      @Increment              CHAR(1),
      @StartDate              DATETIME,
      @EndDate                DATETIME
)
RETURNS  
@SelectedRange    TABLE 
(IndividualDate DATETIME)
AS 
BEGIN
      WITH cteRange (DateRange) AS (
            SELECT DATEADD(dd, 0, DATEDIFF(dd, 0, @StartDate))
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
          
      INSERT INTO @SelectedRange (IndividualDate)
      SELECT DateRange
      FROM cteRange
      OPTION (MAXRECURSION 3660);
      RETURN
END
GO