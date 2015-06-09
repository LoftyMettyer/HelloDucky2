/* -------------------------------------- */
/* CLUK                                   */
/* Populate Salary_increase_or_not column */
/* Mark Edwynn - June 2015                */
/* -------------------------------------- */

CREATE TABLE #IncreaseNot
(
      [ID] integer
	, [Increase_Not] varchar(16)
);

INSERT INTO #IncreaseNot ([ID], [Increase_Not])
SELECT [ID], CASE WHEN [Amount] = (SELECT TOP 1 [Amount] FROM [dbo].[tbuser_Salary] WHERE [iD_1] = s.[ID_1] AND [Salary_Date] < s.[Salary_Date] ORDER BY [ID_1], [Salary_Date] DESC) THEN 'No Increase' 
		 WHEN [Amount] > (SELECT TOP 1 [Amount] FROM [dbo].[tbuser_Salary] WHERE [iD_1] = s.[ID_1] AND [Salary_Date] < s.[Salary_Date] ORDER BY [ID_1], [Salary_Date] DESC) 
			THEN CASE WHEN [Standard_Hours] != (SELECT TOP 1 [Standard_Hours] FROM [dbo].[tbuser_Salary] WHERE [iD_1] = s.[ID_1] AND [Salary_Date] < s.[Salary_Date] ORDER BY [ID_1], [Salary_Date] DESC)
				THEN 'Increase (hours)' 
				ELSE 'Increase'
			END
		 WHEN [Amount] < (SELECT TOP 1 [Amount] FROM [dbo].[tbuser_Salary] WHERE [iD_1] = s.[ID_1] AND [Salary_Date] < s.[Salary_Date] ORDER BY [ID_1], [Salary_Date] DESC)
			THEN CASE WHEN [Standard_Hours] != (SELECT TOP 1 [Standard_Hours] FROM [dbo].[tbuser_Salary] WHERE [iD_1] = s.[ID_1] AND [Salary_Date] < s.[Salary_Date] ORDER BY [ID_1], [Salary_Date] DESC)
				THEN 'Decrease (hours)'
				ELSE 'Decrease'
			END
		 ELSE 'New Starter'
	END AS Increase_Not
FROM [dbo].[tbuser_Salary] s
ORDER BY [ID];

DISABLE TRIGGER ALL ON [dbo].[tbuser_Salary];

UPDATE [dbo].[tbuser_Salary]
SET [Salary_increase_or_not] = base.[Increase_Not]
FROM #IncreaseNot base WHERE base.[ID] = [dbo].[tbuser_Salary].[ID];

ENABLE TRIGGER ALL ON [dbo].[tbuser_Salary];

DROP TABLE #IncreaseNot;
