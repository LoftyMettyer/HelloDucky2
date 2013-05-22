CREATE PROCEDURE [dbo].[sp_ASRFn_NiceTime]
(
	@psResult 		varchar(MAX) OUTPUT,
	@psTimeString	varchar(MAX) -- in the format hh:mm:ss (24 hour clock)
)
AS
BEGIN

	-- Return the given time in the format hh:mm am/pm (12 hour clock)
	select @psResult = 
	case 
		when len(ltrim(rtrim(@psTimeString))) = 0 then ''
		else 
			case 
				when isdate(@psTimeString) = 0 then '***'
				else (convert(varchar(2),((datepart(hour,convert(datetime, @psTimeString)) + 11) % 12) + 1)
					+ ':' + right('00' + datename(minute, convert(datetime, @psTimeString)),2)
					+ case 
						when datepart(hour, convert(datetime, @psTimeString)) > 11 then ' pm'
						else ' am' 
					end) 
			end 
	end
END
