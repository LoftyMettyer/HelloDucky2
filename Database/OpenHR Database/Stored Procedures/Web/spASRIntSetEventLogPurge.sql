CREATE PROCEDURE [dbo].[spASRIntSetEventLogPurge]
(
		@psPeriod		varchar(2),
		@piFrequency	integer
)
AS
BEGIN

	SET NOCOUNT ON;

	INSERT INTO [dbo].[ASRSysEventLogPurge] (Period,Frequency)
	VALUES (@psPeriod, @piFrequency);

END