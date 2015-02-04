CREATE  FUNCTION [dbo].[udfsysDurationFromPattern](
		@Absence_In	varchar(5),
		@IndividualDate datetime,
		@SessionType varchar(3),
		@Sunday_Hours_AM numeric(4,2),
		@Monday_Hours_AM numeric(4,2),
		@Tuesday_Hours_AM numeric(4,2),
		@Wednesday_Hours_AM numeric(4,2),
		@Thursday_Hours_AM numeric(4,2),
		@Friday_Hours_AM numeric(4,2),
		@Saturday_Hours_AM numeric(4,2),
		@Sunday_Hours_PM numeric(4,2),
		@Monday_Hours_PM numeric(4,2),
		@Tuesday_Hours_PM numeric(4,2),
		@Wednesday_Hours_PM numeric(4,2),
		@Thursday_Hours_PM numeric(4,2),
		@Friday_Hours_PM numeric(4,2),
		@Saturday_Hours_PM numeric(4,2))
	RETURNS numeric(5,2)
	AS 
	BEGIN

		DECLARE @value numeric(5,2) = 0;

		IF @Absence_In = 'Days'
		BEGIN
			IF @Sunday_Hours_AM > 0 SET @Sunday_Hours_AM = 0.5
			IF @Monday_Hours_AM > 0 SET @Monday_Hours_AM = 0.5
			IF @Tuesday_Hours_AM > 0 SET @Tuesday_Hours_AM = 0.5
			IF @Wednesday_Hours_AM > 0 SET @Wednesday_Hours_AM = 0.5
			IF @Thursday_Hours_AM > 0 SET @Thursday_Hours_AM = 0.5
			IF @Friday_Hours_AM > 0 SET @Friday_Hours_AM = 0.5
			IF @Saturday_Hours_AM > 0 SET @Saturday_Hours_AM = 0.5
			IF @Sunday_Hours_PM > 0 SET @Sunday_Hours_PM = 0.5
			IF @Monday_Hours_PM > 0 SET @Monday_Hours_PM = 0.5
			IF @Tuesday_Hours_PM > 0 SET @Tuesday_Hours_PM = 0.5
			IF @Wednesday_Hours_PM > 0 SET @Wednesday_Hours_PM = 0.5
			IF @Thursday_Hours_PM > 0 SET @Thursday_Hours_PM = 0.5
			IF @Friday_Hours_PM > 0 SET @Friday_Hours_PM = 0.5
			IF @Saturday_Hours_PM > 0 SET @Saturday_Hours_PM = 0.5
		END

		SET @value = ISNULL(CASE WHEN DATEPART(dw, @IndividualDate) = 1 AND @SessionType = 'AM' THEN @Sunday_Hours_AM
				WHEN DATEPART(dw, @IndividualDate) = 2 AND @SessionType = 'AM' THEN @Monday_Hours_AM
				WHEN DATEPART(dw, @IndividualDate) = 3 AND @SessionType = 'AM' THEN @Tuesday_Hours_AM
				WHEN DATEPART(dw, @IndividualDate) = 4 AND @SessionType = 'AM' THEN @Wednesday_Hours_AM
				WHEN DATEPART(dw, @IndividualDate) = 5 AND @SessionType = 'AM' THEN @Thursday_Hours_AM
				WHEN DATEPART(dw, @IndividualDate) = 6 AND @SessionType = 'AM' THEN @Friday_Hours_AM
				WHEN DATEPART(dw, @IndividualDate) = 7 AND @SessionType = 'AM' THEN @Saturday_Hours_AM
				WHEN DATEPART(dw, @IndividualDate) = 1 AND @SessionType = 'PM' THEN @Sunday_Hours_PM
				WHEN DATEPART(dw, @IndividualDate) = 2 AND @SessionType = 'PM' THEN @Monday_Hours_PM
				WHEN DATEPART(dw, @IndividualDate) = 3 AND @SessionType = 'PM' THEN @Tuesday_Hours_PM
				WHEN DATEPART(dw, @IndividualDate) = 4 AND @SessionType = 'PM' THEN @Wednesday_Hours_PM
				WHEN DATEPART(dw, @IndividualDate) = 5 AND @SessionType = 'PM' THEN @Thursday_Hours_PM
				WHEN DATEPART(dw, @IndividualDate) = 6 AND @SessionType = 'PM' THEN @Friday_Hours_PM
				WHEN DATEPART(dw, @IndividualDate) = 7 AND @SessionType = 'PM' THEN @Saturday_Hours_PM
				WHEN DATEPART(dw, @IndividualDate) = 1 AND @SessionType = 'Day' THEN @Sunday_Hours_AM + @Sunday_Hours_PM
				WHEN DATEPART(dw, @IndividualDate) = 2 AND @SessionType = 'Day' THEN @Monday_Hours_AM + @Monday_Hours_PM
				WHEN DATEPART(dw, @IndividualDate) = 3 AND @SessionType = 'Day' THEN @Tuesday_Hours_AM + @Tuesday_Hours_PM
				WHEN DATEPART(dw, @IndividualDate) = 4 AND @SessionType = 'Day' THEN @Wednesday_Hours_AM + @Wednesday_Hours_PM
				WHEN DATEPART(dw, @IndividualDate) = 5 AND @SessionType = 'Day' THEN @Thursday_Hours_AM + @Thursday_Hours_PM
				WHEN DATEPART(dw, @IndividualDate) = 6 AND @SessionType = 'Day' THEN @Friday_Hours_AM + @Friday_Hours_PM
				WHEN DATEPART(dw, @IndividualDate) = 7 AND @SessionType = 'Day' THEN @Saturday_Hours_AM + @Saturday_Hours_PM
			END, 0)

		RETURN @value;

	END






