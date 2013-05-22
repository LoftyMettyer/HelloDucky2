CREATE FUNCTION [dbo].[udfASRAddWeekdays]
(
	@StartDate datetime, 
	@Duration int
)
RETURNS datetime
AS
BEGIN

	DECLARE @ReturnDate datetime

	IF NULLIF(@Duration, 0) IS NULL	
		RETURN @StartDate

	SELECT @ReturnDate = DATEADD(d,
						CASE DATEPART(dw,@StartDate) 
						WHEN 7 THEN 2 
						WHEN 1 THEN 1 
						ELSE 0 END,	@StartDate)
						+(DATEPART(dw,DATEADD(d,
							CASE DATEPART(dw,@StartDate) 
							WHEN 7 THEN 2 
							WHEN 1 THEN 1 
							ELSE 0 END,@StartDate))-2+@Duration)%5
						+((DATEPART(dw,DATEADD(d,
							CASE DATEPART(dw,@StartDate) 
							WHEN 7 THEN 2 
							WHEN 1 THEN 1 
							ELSE 0 END,@StartDate))-2+@Duration)/5)*7
						-(DATEPART(dw,DATEADD(d,
							CASE DATEPART(dw,@StartDate) 
							WHEN 7 THEN 2 
							WHEN 1 THEN 1 
							ELSE 0 END,@StartDate))-2)

	IF @ReturnDate IS NULL
		RETURN @StartDate

	RETURN @ReturnDate
END