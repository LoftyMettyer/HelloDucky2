CREATE PROCEDURE [dbo].[spASRIntSetEventLogPurge]
(
		@psPeriod		varchar(2),
		@piFrequency	integer
)
AS
BEGIN
	INSERT INTO [dbo].[ASRSysEventLogPurge] (Period,Frequency)
	VALUES (@psPeriod, @piFrequency);
END