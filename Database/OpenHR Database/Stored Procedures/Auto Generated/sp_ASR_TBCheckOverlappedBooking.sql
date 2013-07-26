CREATE PROCEDURE [dbo].[sp_ASR_TBCheckOverlappedBooking] (
  @plngCourseRecordID int,
  @plngEmployeeRecordID int,
  @plngBookingRecordID int,
  @piReturnCode int OUTPUT
)
AS
BEGIN
	SET @piReturnCode = 1;
END