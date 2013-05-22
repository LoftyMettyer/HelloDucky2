CREATE PROCEDURE [dbo].[spASRIntGetCalendarColours]
AS
BEGIN
	SELECT ASRSysColours.ColOrder, 
		ASRSysColours.ColValue,
		ASRSysColours.ColDesc, 
		ASRSysColours.WordColourIndex,
		ASRSysColours.CalendarLegendColour
	FROM ASRSysColours
	WHERE (ASRSysColours.CalendarLegendColour = 1)
		AND (ASRSysColours.ColValue NOT IN (13434879))
	ORDER BY ASRSysColours.ColOrder;
END