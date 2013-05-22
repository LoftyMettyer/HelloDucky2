CREATE PROCEDURE sp_ASRFn_IsBetween 
(
    @result bit OUTPUT,
    @datetest datetime,
    @datelower datetime,
    @dateupper datetime,
    @numerictest float,
    @numericlower float,
    @numericupper float
)

AS


if @datetest is not null
begin
	if (@datetest >= @datelower) and (@datetest <= @dateupper)
	begin
		set @result = 1
	end
	if not (@datetest >= @datelower) or not (@datetest <= @dateupper)
	begin
		set @result = 0
	end
end

if @numerictest is not null
begin
	if (@numerictest >= @numericlower) and (@numerictest <= @numericupper)
	begin
		set @result = 1
	end
	if not (@numerictest >= @numericlower) or not (@numerictest <= @numericupper)
	begin
		set @result = 0
	end
end

select @result as result

GO

